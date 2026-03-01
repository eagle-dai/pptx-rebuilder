import os
import io
import re
import base64
import json
import zipfile
from typing import Optional
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from langchain_core.messages import HumanMessage
from PIL import Image

# 引入强大的 EasyOCR
import easyocr

# SAP Cloud SDK for AI
from gen_ai_hub.proxy.langchain.init_models import init_llm

app = FastAPI(title="PPTX Rebuilder API (EasyOCR + LLM)")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 全局初始化 EasyOCR Reader，避免每次请求重复加载模型
# 首次运行会自动下载模型到本地 (约 10-20MB)
print("正在加载 EasyOCR 模型 (中英文)...")
ocr_reader = easyocr.Reader(["en", "ch_sim"])
print("EasyOCR 模型加载完成！")


def fix_json_quotes(json_str: str) -> str:
    json_str = json_str.replace(""", "'").replace(""", "'")
    json_str = json_str.replace('"', '"').replace('"', '"')
    result = []
    in_string = False
    i = 0
    while i < len(json_str):
        c = json_str[i]
        if c == "\\" and i + 1 < len(json_str):
            result.append(c)
            result.append(json_str[i + 1])
            i += 2
            continue
        if c == '"':
            if not in_string:
                in_string = True
                result.append(c)
            else:
                j = i + 1
                while j < len(json_str) and json_str[j] in " \t\n\r":
                    j += 1
                if j >= len(json_str) or json_str[j] in ",}]:":
                    in_string = False
                    result.append(c)
                else:
                    result.append("'")
        else:
            result.append(c)
        i += 1
    return "".join(result)


def get_llm():
    model_name = os.getenv("LLM_MODEL_NAME", "anthropic--claude-4.5-sonnet")
    llm = init_llm(model_name, max_tokens=8192)
    if hasattr(llm, "top_p"):
        llm.top_p = None
    return llm


def extract_slide_images_from_pptx(
    pptx_bytes: io.BytesIO,
) -> list[tuple[int, str, bytes]]:
    images = []
    pptx_bytes.seek(0)
    with zipfile.ZipFile(pptx_bytes, "r") as zf:
        slide_files = sorted(
            [
                f
                for f in zf.namelist()
                if f.startswith("ppt/slides/slide") and f.endswith(".xml")
            ],
            key=lambda x: int(x.replace("ppt/slides/slide", "").replace(".xml", "")),
        )

        for slide_idx, slide_file in enumerate(slide_files):
            slide_num = slide_file.replace("ppt/slides/slide", "").replace(".xml", "")
            rels_file = f"ppt/slides/_rels/slide{slide_num}.xml.rels"
            if rels_file not in zf.namelist():
                continue
            rels_content = zf.read(rels_file).decode("utf-8")
            image_refs = re.findall(
                r'Target="\.\./(media/image\d+\.[a-zA-Z]+)"', rels_content
            )

            if image_refs:
                image_path = f"ppt/{image_refs[0]}"
                if image_path in zf.namelist():
                    image_data = zf.read(image_path)
                    img = Image.open(io.BytesIO(image_data))
                    png_buffer = io.BytesIO()
                    img.convert("RGB").save(png_buffer, format="PNG")
                    png_buffer.seek(0)
                    main_image_base64 = base64.b64encode(png_buffer.read()).decode(
                        "utf-8"
                    )
                    images.append((slide_idx, main_image_base64, image_data))
    return images


def perform_ocr(image_bytes: bytes) -> list[dict]:
    """
    使用 EasyOCR 提取图像中的文本行及其精确坐标。
    """
    img = Image.open(io.BytesIO(image_bytes))
    w, h = img.size

    # EasyOCR 直接处理字节流
    results = ocr_reader.readtext(image_bytes)

    ocr_data = []
    for bbox, text, prob in results:
        # 过滤低置信度和空白内容
        if prob > 0.3 and text.strip():
            # bbox 是四个角点的坐标 [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
            x_coords = [p[0] for p in bbox]
            y_coords = [p[1] for p in bbox]

            left = min(x_coords)
            right = max(x_coords)
            top = min(y_coords)
            bottom = max(y_coords)

            # 转换为百分比坐标，方便后续 LLM 处理
            ocr_data.append(
                {
                    "text": text,
                    "left": f"{int(left / w * 100)}%",
                    "top": f"{int(top / h * 100)}%",
                    "width": f"{int((right - left) / w * 100)}%",
                    "height": f"{int((bottom - top) / h * 100)}%",
                }
            )

    return ocr_data


def parse_color(color_str: str) -> Optional[RGBColor]:
    if not color_str:
        return None
    color_str = color_str.strip().lower()

    semantic_colors = {
        "blue": RGBColor(0, 112, 242),
        "sap blue": RGBColor(0, 112, 242),
        "navy": RGBColor(10, 30, 80),
        "black": RGBColor(0, 0, 0),
        "white": RGBColor(255, 255, 255),
        "gray": RGBColor(128, 128, 128),
        "grey": RGBColor(128, 128, 128),
        "light gray": RGBColor(240, 240, 240),
        "red": RGBColor(230, 50, 50),
        "green": RGBColor(40, 160, 60),
    }
    if color_str in semantic_colors:
        return semantic_colors[color_str]

    if color_str.startswith("#"):
        hex_color = color_str[1:]
        if len(hex_color) == 3:
            hex_color = "".join([c * 2 for c in hex_color])
        if len(hex_color) == 6:
            try:
                return RGBColor(
                    int(hex_color[0:2], 16),
                    int(hex_color[2:4], 16),
                    int(hex_color[4:6], 16),
                )
            except ValueError:
                pass
    return None


def create_slide_from_analysis(prs: Presentation, analysis: dict):
    try:
        blank_layout = prs.slide_layouts[6]
    except IndexError:
        blank_layout = prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank_layout)

    slide_width, slide_height = prs.slide_width, prs.slide_height

    bg_color = analysis.get("background_color")
    if bg_color:
        parsed_color = parse_color(bg_color)
        if parsed_color:
            background = slide.background
            background.fill.solid()
            background.fill.fore_color.rgb = parsed_color

    elements = analysis.get("elements", [])
    elements.sort(key=lambda x: x.get("z_order", 10))

    for element in elements:
        elem_type = element.get("type", "text")
        left = _parse_position(element.get("left", 0), slide_width)
        top = _parse_position(element.get("top", 0), slide_height)
        width = _parse_position(element.get("width", 100), slide_width)
        height = _parse_position(element.get("height", 10), slide_height)

        if elem_type == "text":
            _add_text_element(slide, element, left, top, width, height)
        elif elem_type == "image_placeholder":
            _add_image_placeholder(slide, element, left, top, width, height)
        elif elem_type == "table":
            _add_table_element(slide, element, left, top, width, height)
        elif elem_type == "shape":
            _add_shape_element(slide, element, left, top, width, height)
    return slide


def _parse_position(value, reference_size) -> int:
    """极其强壮的位置解析器，专治各种 LLM 幻觉"""
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        value = value.strip()
        # 处理正常或异常的百分比 (例如 "50%", "50%50%", "50")
        if "%" in value:
            match = re.search(r"(\d+(?:\.\d+)?)", value)
            if match:
                return int(reference_size * (float(match.group(1)) / 100))
        # 处理带有 px 或是疯狂重复的字符串 (例如 "32px32px32px")
        match = re.search(r"(\d+(?:\.\d+)?)", value)
        if match:
            # 提取第一个数字，按近似像素转换为 PPT 内部的 EMU 单位
            return int(float(match.group(1)) * 914400 / 96)
    return 0


def _parse_safe_number(value, default=14) -> int:
    """极其强壮的数字提取器（用于字号等）"""
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        match = re.search(r"(\d+)", value)
        if match:
            return int(match.group(1))
    return default


def _add_text_element(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    """添加文本元素，带有字号容错处理"""
    content = element.get("content", "")
    if not content:
        return
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    for i, line in enumerate(content.split("\n")):
        para = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        para.text = line

        # 使用安全的数字提取器防止字号崩溃
        safe_font_size = _parse_safe_number(element.get("font_size", 14), default=14)
        para.font.size = Pt(safe_font_size)

        para.font.name = element.get("font_name", "Microsoft YaHei")
        if element.get("bold"):
            para.font.bold = True

        color = element.get("color")
        if color:
            parsed = parse_color(color)
            if parsed:
                para.font.color.rgb = parsed

        align = element.get("align", "left")
        if align == "center":
            para.alignment = PP_ALIGN.CENTER
        elif align == "right":
            para.alignment = PP_ALIGN.RIGHT
        else:
            para.alignment = PP_ALIGN.LEFT


def _add_image_placeholder(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(240, 245, 250)
    shape.line.color.rgb = RGBColor(200, 210, 220)
    text_frame = shape.text_frame
    text_frame.word_wrap = True
    para = text_frame.paragraphs[0]
    para.text = f"[插图区]\n{element.get('description', '插图占位')}"
    para.alignment = PP_ALIGN.CENTER
    para.font.size = Pt(12)
    para.font.color.rgb = RGBColor(120, 130, 140)
    para.font.name = "Microsoft YaHei"


def _add_shape_element(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    mso_shape = (
        MSO_SHAPE.ROUNDED_RECTANGLE
        if element.get("shape_type") == "rounded_rectangle"
        else MSO_SHAPE.RECTANGLE
    )
    shape = slide.shapes.add_shape(mso_shape, left, top, width, height)

    fill_color = element.get("fill_color")
    if fill_color:
        parsed = parse_color(fill_color)
        if parsed:
            shape.fill.solid()
            shape.fill.fore_color.rgb = parsed
        else:
            shape.fill.background()
    else:
        shape.fill.background()

    border_color = element.get("border_color")
    if border_color:
        parsed_border = parse_color(border_color)
        if parsed_border:
            shape.line.color.rgb = parsed_border
    else:
        shape.line.fill.background()


def _add_table_element(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    rows_data = element.get("rows", [])
    if not rows_data:
        return
    num_rows, num_cols = (
        len(rows_data),
        max(len(r) for r in rows_data) if rows_data else 1,
    )
    table = slide.shapes.add_table(num_rows, num_cols, left, top, width, height).table

    for r_idx, row in enumerate(rows_data):
        for c_idx, cell_text in enumerate(row):
            if c_idx < num_cols:
                cell = table.cell(r_idx, c_idx)
                cell.text = str(cell_text)
                cell.text_frame.word_wrap = True
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(12)
                para.font.name = "Microsoft YaHei"
                if element.get("color"):
                    p_c = parse_color(element.get("color"))
                    if p_c:
                        para.font.color.rgb = p_c
                if r_idx == 0 and element.get("has_header", True):
                    para.font.bold = True


def analyze_slide_image_with_ocr(llm, image_base64: str, ocr_data: list) -> dict:
    """
    将原生 OCR 数据和原图一起发给 Claude，让其作为“清洗与合并引擎”。
    """
    ocr_json_str = json.dumps(ocr_data, ensure_ascii=False)

    message = HumanMessage(
        content=[
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": image_base64,
                },
            },
            {
                "type": "text",
                "text": f"""你是一个高级版面重构算法。我已经通过专业的 OCR 引擎提取了这张图片的精确物理坐标和文字，数据如下：
<ocr_raw_data>
{ocr_json_str}
</ocr_raw_data>

你的任务是：基于上面的 OCR 坐标数据和提供的原图，进行【语义清洗与元素聚合】。

【极其严格的规则】：
1. **聚合重组**：OCR 数据是逐行扫描的，非常零碎。你必须观察原图，如果某几行 OCR 数据在视觉上同属于一个卡片、一个段落或一个表格单元格，你【必须】将它们合并为一个 `type: "text"` 元素（使用 `\\n` 换行）。
2. **继承坐标**：合并后的文本块的 `left/top/width/height`，必须基于组成它的 OCR 子块的坐标域进行外推包含（涵盖它们的极值边界）。绝不准自己瞎编坐标！
3. **表格构建**：如果 OCR 数据在视觉上呈现网格排列，你必须生成 `type: "table"`，把对应的 OCR 文字填入 `rows`，并计算整个表格的边界坐标。
4. **补充色彩和背景**：OCR 引擎没有颜色信息。你需要看原图，为合并后的文本补充 `"color"` 字段，并根据需要为它们底层垫上 `"shape"` 背景块（z_order: 0）。
5. **复杂图形占位**：对 OCR 无法提取的复杂架构图、流程图区域，生成 `image_placeholder`。

【输出要求】：
仅返回包含 `elements` 数组的纯 JSON。不需要解释。""",
            },
        ]
    )

    response = llm.invoke([message])

    try:
        response_text = response.content.strip()
        code_block_match = re.search(
            r"```(?:json)?\s*\n?(.*?)\n?```", response_text, re.DOTALL
        )
        if code_block_match:
            response_text = code_block_match.group(1).strip()
        if response_text.startswith("```"):
            lines = response_text.split("\n")
            response_text = "\n".join(lines[1:-1]).strip()

        json_start = response_text.find("{")
        json_end = response_text.rfind("}")
        if json_start != -1 and json_end != -1:
            response_text = response_text[json_start : json_end + 1]

        response_text = fix_json_quotes(response_text)
        return json.loads(response_text)

    except Exception as e:
        print(f"JSON 解析失败: {e}")
        return {
            "elements": [
                {
                    "type": "text",
                    "content": "大模型处理 OCR 数据失败",
                    "left": "10%",
                    "top": "10%",
                    "width": "80%",
                    "height": "10%",
                }
            ]
        }


@app.post("/api/convert")
async def convert_pptx(file: UploadFile = File(...)):
    if not file.filename.endswith(".pptx"):
        raise HTTPException(status_code=400, detail="仅支持 .pptx 文件上传")
    try:
        content = await file.read()
        source_pptx = io.BytesIO(content)
        slide_images = extract_slide_images_from_pptx(source_pptx)

        if not slide_images:
            raise HTTPException(status_code=400, detail="未能提取到图片。")

        llm = get_llm()
        target_presentation = Presentation()
        target_presentation.slide_width = Inches(13.333)
        target_presentation.slide_height = Inches(7.5)

        for slide_idx, image_base64, image_raw_bytes in slide_images:
            # 1. 使用 EasyOCR 进行像素级坐标与文字提取
            print(f"正在对第 {slide_idx + 1} 页进行 OCR 物理扫描...")
            ocr_data = perform_ocr(image_raw_bytes)

            # 2. 交给 Claude 进行语义组合与颜色提取
            print(f"正在交给 LLM 进行语义排版清洗...")
            analysis = analyze_slide_image_with_ocr(llm, image_base64, ocr_data)

            # 3. 渲染
            create_slide_from_analysis(target_presentation, analysis)

        output = io.BytesIO()
        target_presentation.save(output)
        output.seek(0)
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f"attachment; filename=rebuilt_EasyOCR_{file.filename}"
            },
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"转换失败: {str(e)}")


@app.get("/api/health")
def health_check():
    return {"status": "ok", "message": "Backend is running with EasyOCR Architecture."}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
