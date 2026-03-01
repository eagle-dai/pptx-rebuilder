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

# 全局初始化 EasyOCR Reader
print("正在加载 EasyOCR 模型 (中英文)...")
ocr_reader = easyocr.Reader(["en", "ch_sim"])
print("EasyOCR 模型加载完成！")


def fix_json_quotes(json_str: str) -> str:
    json_str = json_str.replace("“", '"').replace("”", '"')
    json_str = json_str.replace("‘", "'").replace("’", "'")
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
    img = Image.open(io.BytesIO(image_bytes))
    w, h = img.size

    # 放宽合并策略，让 EasyOCR 尝试合并相近的段落
    results = ocr_reader.readtext(image_bytes, paragraph=False)

    ocr_data = []
    for bbox, text, prob in results:
        if prob > 0.3 and text.strip():
            x_coords = [p[0] for p in bbox]
            y_coords = [p[1] for p in bbox]

            left = min(x_coords)
            right = max(x_coords)
            top = min(y_coords)
            bottom = max(y_coords)

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
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        value = value.strip()
        if "%" in value:
            match = re.search(r"(\d+(?:\.\d+)?)", value)
            if match:
                return int(reference_size * (float(match.group(1)) / 100))
        match = re.search(r"(\d+(?:\.\d+)?)", value)
        if match:
            return int(float(match.group(1)) * 914400 / 96)
    return 0


def _parse_safe_number(value, default=14) -> int:
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

        safe_font_size = _parse_safe_number(element.get("font_size", 14), default=14)
        para.font.size = Pt(safe_font_size)
        para.font.name = element.get("font_name", "Microsoft YaHei")

        # 处理加粗标识
        if element.get("bold") is True:
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
    if not rows_data or not isinstance(rows_data, list):
        return

    # 防止空行崩溃
    valid_rows = [row for row in rows_data if isinstance(row, list) and len(row) > 0]
    if not valid_rows:
        return

    num_rows = len(valid_rows)
    num_cols = max(len(r) for r in valid_rows)
    table = slide.shapes.add_table(num_rows, num_cols, left, top, width, height).table

    for r_idx, row in enumerate(valid_rows):
        for c_idx, cell_text in enumerate(row):
            if c_idx < num_cols:
                cell = table.cell(r_idx, c_idx)
                cell.text = str(cell_text) if cell_text else ""
                cell.text_frame.word_wrap = True
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(12)
                para.font.name = "Microsoft YaHei"

                # 设置表头加粗
                if r_idx == 0 and element.get("has_header", True):
                    para.font.bold = True
                # 全局表格字体颜色
                if element.get("color"):
                    p_c = parse_color(element.get("color"))
                    if p_c:
                        para.font.color.rgb = p_c


def analyze_slide_image_with_ocr(llm, image_base64: str, ocr_data: list) -> dict:
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
                "text": f"""你是一个顶尖的文档版面重构专家（Document Layout Analyst）。我已经通过专业的 OCR 引擎提取了这张图片的物理坐标和文字，数据如下：
<ocr_raw_data>
{ocr_json_str}
</ocr_raw_data>

请基于上面的 OCR 坐标数据和提供的原图，进行【深度语义清洗与元素聚合】。

【绝对禁令与严格规则】：
1. **彻底消灭占位符**：绝对不允许在表格或文本中使用 "cells", "text here" 等无意义的占位符！所有内容必须 100% 来源于 OCR 原文提取。
2. **强制语义聚合**：原图中的标题、长段落或形状内部的文字，在 OCR 中通常是断行的。你【必须】将视觉上同属一个区块（例如同一个箭头、同一个文本框）的碎片化 OCR 文本合并为一个 `type: "text"` 元素，使用 `\\n` 换行。
3. **精准表格构建**：若画面存在表格网格结构，必须使用 `type: "table"`。`rows` 必须是一个严格的二维数组，将对应的 OCR 文本完美填入每一个单元格中。如果单元格包含多行文字，也必须完整保留。
4. **字体层级识别**：根据原图中文字的大小与粗细，主动为 `type: "text"` 补充 `"font_size"` (例如大标题用 32 或更大，正文用 14) 和 `"bold"` (布尔值) 属性。

【必须遵循的 JSON 输出 Schema 示例】：
```json
{{
  "background_color": "#FFFFFF",
  "elements": [
    {{
      "type": "text",
      "content": "大标题合并\\n第二行副标题",
      "left": "5%", "top": "5%", "width": "50%", "height": "10%",
      "font_size": 36, "bold": true, "color": "#0B2D71", "align": "left", "z_order": 10
    }},
    {{
      "type": "shape",
      "shape_type": "rectangle",
      "fill_color": "#0070F2",
      "left": "70%", "top": "40%", "width": "25%", "height": "20%", "z_order": 1
    }},
    {{
      "type": "table",
      "has_header": true,
      "rows": [
        ["Layer A Latency", "Edge Solution"],
        ["User needs immediate...", "End-device Engine triggers..."]
      ],
      "left": "5%", "top": "25%", "width": "60%", "height": "40%", "z_order": 5
    }},
    {{
      "type": "image_placeholder",
      "description": "复杂的系统架构拓扑图",
      "left": "25%", "top": "15%", "width": "50%", "height": "40%", "z_order": 5
    }}
  ]
}}
```

要求：只返回纯 JSON，严禁输出任何分析过程或 Markdown 之外的文本。""",
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
        print(f"JSON 解析失败: {e}\n模型原文: {response.content}")
        return {
            "elements": [
                {
                    "type": "text",
                    "content": "大模型处理 OCR 数据失败，请检查终端日志。",
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
        # 兼容目前常见的 16:9 宽屏比例
        target_presentation.slide_width = Inches(13.333)
        target_presentation.slide_height = Inches(7.5)

        for slide_idx, image_base64, image_raw_bytes in slide_images:
            print(f"正在对第 {slide_idx + 1} 页进行 OCR 物理扫描...")
            ocr_data = perform_ocr(image_raw_bytes)

            print(f"正在交给 LLM 进行语义排版清洗...")
            analysis = analyze_slide_image_with_ocr(llm, image_base64, ocr_data)

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


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
