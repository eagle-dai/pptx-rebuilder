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
from langchain_core.messages import HumanMessage
from PIL import Image

# SAP Cloud SDK for AI - 使用 LangChain 集成
from gen_ai_hub.proxy.langchain.init_models import init_llm

app = FastAPI(title="PPTX Rebuilder API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


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


def extract_slide_images_from_pptx(pptx_bytes: io.BytesIO) -> list[tuple[int, str]]:
    """
    仅提取每页幻灯片的主图作为 Base64 返回。
    """
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

            main_image_base64 = None

            if image_refs:
                # 默认获取该页面关联的第一张图片作为主图（对于图片型PPT通常就是整页）
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

            if main_image_base64:
                images.append((slide_idx, main_image_base64))

    return images


def parse_color(color_str: str) -> Optional[RGBColor]:
    if not color_str:
        return None

    color_str = color_str.strip().lower()

    color_names = {
        "white": RGBColor(255, 255, 255),
        "black": RGBColor(0, 0, 0),
        "red": RGBColor(255, 0, 0),
        "green": RGBColor(0, 128, 0),
        "blue": RGBColor(0, 0, 255),
        "yellow": RGBColor(255, 255, 0),
        "orange": RGBColor(255, 165, 0),
        "gray": RGBColor(128, 128, 128),
        "grey": RGBColor(128, 128, 128),
        "navy": RGBColor(0, 0, 128),
        "purple": RGBColor(128, 0, 128),
    }

    if color_str in color_names:
        return color_names[color_str]

    if color_str.startswith("#"):
        hex_color = color_str[1:]
        if len(hex_color) == 3:
            hex_color = "".join([c * 2 for c in hex_color])
        if len(hex_color) == 6:
            try:
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                return RGBColor(r, g, b)
            except ValueError:
                pass

    rgb_match = re.match(r"rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", color_str)
    if rgb_match:
        r, g, b = map(int, rgb_match.groups())
        return RGBColor(min(r, 255), min(g, 255), min(b, 255))

    return None


def create_slide_from_analysis(
    prs: Presentation, analysis: dict, slide_idx: int, image_base64: str
):
    try:
        blank_layout = prs.slide_layouts[6]
    except IndexError:
        blank_layout = prs.slide_layouts[-1]

    slide = prs.slides.add_slide(blank_layout)

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    bg_color = analysis.get("background_color")
    if bg_color:
        parsed_color = parse_color(bg_color)
        if parsed_color:
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = parsed_color

    elements = analysis.get("elements", [])
    elements.sort(key=lambda x: x.get("z_order", 0))

    for element in elements:
        elem_type = element.get("type", "text")

        left = _parse_position(element.get("left", 0), slide_width)
        top = _parse_position(element.get("top", 0), slide_height)
        width = _parse_position(element.get("width", 100), slide_width)
        height = _parse_position(element.get("height", 50), slide_height)

        if elem_type == "text":
            _add_text_element(slide, element, left, top, width, height)
        elif elem_type == "image":
            # 使用动态裁剪技术
            _add_image_element_by_cropping(
                slide, image_base64, left, top, width, height, slide_width, slide_height
            )
        elif elem_type == "shape":
            _add_shape_element(slide, element, left, top, width, height)
        elif elem_type == "table":
            _add_table_element(slide, element, left, top, width, height)

    return slide


def _parse_position(value, reference_size) -> int:
    if isinstance(value, str):
        value = value.strip()
        if value.endswith("%"):
            pct = float(value[:-1]) / 100
            return int(reference_size * pct)
        elif value.endswith("px"):
            px = float(value[:-2])
            return int(px * 914400 / 96)
    return int(value) if value else 0


def _add_text_element(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    content = element.get("content", "")
    if not content:
        return

    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame

    # 开启文本自动换行和自适应框体大小，防止溢出
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    v_align = element.get("vertical_align", "top")
    if v_align == "middle":
        text_frame.anchor = MSO_ANCHOR.MIDDLE
    elif v_align == "bottom":
        text_frame.anchor = MSO_ANCHOR.BOTTOM
    else:
        text_frame.anchor = MSO_ANCHOR.TOP

    lines = content.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            para = text_frame.paragraphs[0]
        else:
            para = text_frame.add_paragraph()

        para.text = line

        font_size = element.get("font_size", 18)
        para.font.size = Pt(font_size)
        para.font.name = element.get("font_name", "Microsoft YaHei")

        if element.get("bold"):
            para.font.bold = True
        if element.get("italic"):
            para.font.italic = True

        color = element.get("color")
        if color:
            parsed_color = parse_color(color)
            if parsed_color:
                para.font.color.rgb = parsed_color

        align = element.get("align", "left")
        if align == "center":
            para.alignment = PP_ALIGN.CENTER
        elif align == "right":
            para.alignment = PP_ALIGN.RIGHT
        else:
            para.alignment = PP_ALIGN.LEFT


def _add_image_element_by_cropping(
    slide,
    image_base64: str,
    left: int,
    top: int,
    width: int,
    height: int,
    slide_width: int,
    slide_height: int,
):
    """根据 LLM 给出的坐标，从原图中动态裁剪出对应区域并插入。"""
    try:
        image_data = base64.b64decode(image_base64)
        img = Image.open(io.BytesIO(image_data))

        img_w, img_h = img.size

        # 将 EMU 坐标转换回图片的像素坐标比例
        crop_left = int((left / slide_width) * img_w)
        crop_top = int((top / slide_height) * img_h)
        crop_right = int(((left + width) / slide_width) * img_w)
        crop_bottom = int(((top + height) / slide_height) * img_h)

        crop_left = max(0, crop_left)
        crop_top = max(0, crop_top)
        crop_right = min(img_w, crop_right)
        crop_bottom = min(img_h, crop_bottom)

        if crop_right > crop_left and crop_bottom > crop_top:
            cropped_img = img.crop((crop_left, crop_top, crop_right, crop_bottom))
            img_byte_arr = io.BytesIO()
            cropped_img.save(img_byte_arr, format="PNG")
            img_byte_arr.seek(0)

            slide.shapes.add_picture(img_byte_arr, left, top, width, height)
    except Exception as e:
        print(f"动态裁剪图片失败: {e}")


def _add_shape_element(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    from pptx.enum.shapes import MSO_SHAPE

    shape_type = element.get("shape_type", "rectangle")
    shape_map = {
        "rectangle": MSO_SHAPE.RECTANGLE,
        "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
        "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    }

    mso_shape = shape_map.get(shape_type, MSO_SHAPE.RECTANGLE)
    shape = slide.shapes.add_shape(mso_shape, left, top, width, height)

    fill_color = element.get("fill_color")
    if fill_color:
        parsed_color = parse_color(fill_color)
        if parsed_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = parsed_color

    line_color = element.get("line_color")
    if line_color:
        parsed_color = parse_color(line_color)
        if parsed_color:
            shape.line.color.rgb = parsed_color


def _add_table_element(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    rows_data = element.get("rows", [])
    if not rows_data:
        return

    num_rows = len(rows_data)
    num_cols = max(len(row) for row in rows_data) if rows_data else 1

    table = slide.shapes.add_table(num_rows, num_cols, left, top, width, height).table

    for row_idx, row_data in enumerate(rows_data):
        for col_idx, cell_text in enumerate(row_data):
            if col_idx < num_cols:
                cell = table.cell(row_idx, col_idx)
                cell.text = str(cell_text)

                # 为表格内文字增加自适应，尽量避免越界
                text_frame = cell.text_frame
                text_frame.word_wrap = True

                if row_idx == 0 and element.get("has_header", True):
                    para = cell.text_frame.paragraphs[0]
                    para.font.bold = True
                    para.font.size = Pt(14)


def analyze_slide_image(llm, image_base64: str, slide_index: int) -> dict:
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
                "text": """请仔细分析这张幻灯片图片，提取所有可见元素的内容、位置和样式，返回 JSON 格式。

核心要求：
1. 对于结构复杂的表格（如合并单元格、列宽不一），请放弃使用 table 类型，转而使用多个 type: "text" 元素进行精确定位。
2. 对于复杂的中心图表、流程图、复杂的装饰箭头、人物头像等无法用标准形状绘制的区域，请将其指定为 type: "image"。代码将会根据你给出的 left/top/width/height 坐标系从原图中精准裁剪出该部分并插入。位置评估请尽量精确。

返回格式：
{
  "background_color": "背景色，如 #FFFFFF 或 white（如果是纯色）",
  "elements": [
    {
      "type": "text",
      "content": "文本内容（保留换行符 \\n）",
      "left": "距左边距离（百分比，如 5%）",
      "top": "距顶部距离（百分比，如 10%）",
      "width": "宽度（百分比，如 90%）",
      "height": "高度（百分比，如 15%）",
      "font_size": 32,
      "bold": true,
      "italic": false,
      "color": "文字颜色，如 #333333 或 white",
      "align": "left/center/right",
      "vertical_align": "top/middle/bottom",
      "z_order": 1
    },
    {
      "type": "image",
      "description": "复杂的图表或箭头",
      "left": "25%",
      "top": "20%",
      "width": "50%",
      "height": "40%",
      "z_order": 2
    }
  ]
}

重要规则：
1. 位置使用百分比，估算元素在幻灯片中的相对位置。
2. z_order 表示层叠顺序，数字越小越在底层。
3. 识别所有可见的文字，并独立建立文本框，不要把大段不同样式的文字塞进同一个框。
4. 只返回纯 JSON，不要用 markdown 代码块。""",
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
            start_idx = 1
            end_idx = len(lines)
            for i, line in enumerate(lines):
                if i > 0 and line.strip() == "```":
                    end_idx = i
                    break
            response_text = "\n".join(lines[start_idx:end_idx]).strip()

        json_start = response_text.find("{")
        json_end = response_text.rfind("}")
        if json_start != -1 and json_end != -1 and json_end > json_start:
            response_text = response_text[json_start : json_end + 1]

        response_text = fix_json_quotes(response_text)
        result = json.loads(response_text)

        if not isinstance(result, dict):
            raise ValueError("返回结果不是字典类型")

        if "title" not in result or not result["title"]:
            result["title"] = f"幻灯片 {slide_index + 1}"

        return result

    except (json.JSONDecodeError, AttributeError, ValueError) as e:
        print(f"JSON 解析失败 (幻灯片 {slide_index + 1}): {e}")
        return {
            "title": f"幻灯片 {slide_index + 1}",
            "elements": [
                {
                    "type": "text",
                    "content": "内容解析失败，请重试",
                    "left": "10%",
                    "top": "10%",
                    "width": "80%",
                    "height": "10%",
                }
            ],
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
            raise HTTPException(
                status_code=400,
                detail="未能从 PPTX 中提取到图片，请确保这是一个图片型 PPT 文件",
            )

        llm = get_llm()
        slide_analyses = []

        for slide_idx, image_base64 in slide_images:
            analysis = analyze_slide_image(llm, image_base64, slide_idx)
            # 传递主图以便后续裁剪
            slide_analyses.append((slide_idx, analysis, image_base64))

        target_presentation = Presentation()
        target_presentation.slide_width = Inches(13.333)
        target_presentation.slide_height = Inches(7.5)

        for slide_idx, analysis, image_base64 in slide_analyses:
            create_slide_from_analysis(
                target_presentation, analysis, slide_idx, image_base64
            )

        output = io.BytesIO()
        target_presentation.save(output)
        output.seek(0)

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f"attachment; filename=rebuilt_{file.filename}"
            },
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"转换失败: {str(e)}")


@app.get("/api/health")
def health_check():
    required_env_vars = [
        "AICORE_CLIENT_ID",
        "AICORE_CLIENT_SECRET",
        "AICORE_AUTH_URL",
        "AICORE_BASE_URL",
        "AICORE_RESOURCE_GROUP",
    ]

    missing_vars = [var for var in required_env_vars if not os.getenv(var)]

    if missing_vars:
        return {
            "status": "warning",
            "message": "Backend is running but SAP AI Core not fully configured",
            "missing_env_vars": missing_vars,
        }

    return {"status": "ok", "message": "Backend is running with SAP AI Core configured"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
