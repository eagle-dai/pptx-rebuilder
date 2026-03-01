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

    return None


def create_slide_from_analysis(prs: Presentation, analysis: dict, slide_idx: int):
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
    # 严格根据 z_order 排序，确保背景形状（如卡片）在底层，文字在上层
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
    if isinstance(value, str):
        value = value.strip()
        if value.endswith("%"):
            pct = float(value[:-1]) / 100
            return int(reference_size * pct)
    return int(value) if value else 0


def _add_text_element(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    content = element.get("content", "")
    if not content:
        return

    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame

    # 开启自适应，防止溢出
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    lines = content.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            para = text_frame.paragraphs[0]
        else:
            para = text_frame.add_paragraph()

        para.text = line

        font_size = element.get("font_size", 16)
        para.font.size = Pt(font_size)
        para.font.name = element.get("font_name", "Microsoft YaHei")

        if element.get("bold"):
            para.font.bold = True

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


def _add_image_placeholder(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    desc = element.get("description", "图片插图")

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)

    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(230, 235, 240)  # 稍微带点蓝灰色的高级占位符
    shape.line.fill.background()

    text_frame = shape.text_frame
    text_frame.word_wrap = True
    para = text_frame.paragraphs[0]
    para.text = f"[插图]\n{desc}"
    para.alignment = PP_ALIGN.CENTER
    para.font.size = Pt(12)
    para.font.color.rgb = RGBColor(120, 130, 140)
    para.font.name = "Microsoft YaHei"


def _add_shape_element(
    slide, element: dict, left: int, top: int, width: int, height: int
):
    shape_type = element.get("shape_type", "rectangle")
    shape_map = {
        "rectangle": MSO_SHAPE.RECTANGLE,
        "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
    }

    mso_shape = shape_map.get(shape_type, MSO_SHAPE.RECTANGLE)
    shape = slide.shapes.add_shape(mso_shape, left, top, width, height)

    fill_color = element.get("fill_color")
    if fill_color:
        parsed_color = parse_color(fill_color)
        if parsed_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = parsed_color
        else:
            shape.fill.background()  # 透明填充
    else:
        shape.fill.background()

    # 处理边框
    border_color = element.get("border_color")
    if border_color:
        parsed_border = parse_color(border_color)
        if parsed_border:
            shape.line.color.rgb = parsed_border
    else:
        shape.line.fill.background()  # 无边框


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

                text_frame = cell.text_frame
                text_frame.word_wrap = True

                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(12)
                para.font.name = "Microsoft YaHei"

                if row_idx == 0 and element.get("has_header", True):
                    para.font.bold = True


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
                "text": """你是一个专业的 PPT 页面重构引擎。你的任务是将这张图片格式的幻灯片，精准地逆向还原为结构化的排版 JSON，以便作为通用工具重绘出可编辑的 PPT。

【核心排版还原规则】（极其重要）：
1. **精准定位**：评估每个视觉块在画面中的 left, top, width, height（使用百分比）。尽可能还原原图的视觉分布，保留留白。
2. **文本分离原则**：
   - 如果是视觉上明显分离的文本块（例如：左右两列的对比、位于不同卡片内的文字、散落在图表周围的标注），**必须拆分为不同的 type: "text" 元素**，各自独立定位。
   - 只有属于同一个逻辑段落，或者同一个紧凑列表（如带有 bullet points 的多行）时，才合并在一个文本元素中，使用 \\n 换行。
3. **几何与卡片还原**：画面中的背景色块、卡片背景、简单的几何框选（如矩形、圆角矩形），请使用 type: "shape" 还原，并提取近似的 fill_color 或 border_color。务必将 shape 的 z_order 设置为较小的值（如 0 或 1），以便垫在文字下方。
4. **表格识别**：如果画面中有明显的行列网格、数据对比矩阵，请使用 type: "table" 进行结构化提取。
5. **复杂图像兜底**：仅针对真实照片、极其复杂的拓扑图、带有不规则人物/连线的插画等**完全无法用基础 PPT 形状拼接**的区域，使用 type: "image_placeholder"。

返回 JSON 格式范例：
{
  "background_color": "#FFFFFF",
  "elements": [
    {
      "type": "shape",
      "shape_type": "rounded_rectangle",
      "left": "5%", "top": "20%", "width": "40%", "height": "60%",
      "fill_color": "#F3F6F9",
      "border_color": "#D1D5DB",
      "z_order": 0
    },
    {
      "type": "text",
      "content": "卡片标题",
      "left": "8%", "top": "22%", "width": "34%", "height": "10%",
      "font_size": 20, "bold": true, "color": "#000000", "align": "left",
      "z_order": 1
    },
    {
      "type": "text",
      "content": "这是一段位于左侧卡片内的正文描述。\\n• 要点一\\n• 要点二",
      "left": "8%", "top": "35%", "width": "34%", "height": "40%",
      "font_size": 14, "bold": false, "color": "#333333", "align": "left",
      "z_order": 1
    },
    {
      "type": "image_placeholder",
      "description": "复杂的神经元网络连接图",
      "left": "50%", "top": "20%", "width": "45%", "height": "60%",
      "z_order": 1
    },
    {
      "type": "table",
      "rows": [["参数名", "数值"], ["延迟", "0.4s"]],
      "has_header": true,
      "left": "10%", "top": "85%", "width": "80%", "height": "10%",
      "z_order": 1
    }
  ]
}

请只返回纯 JSON 内容，不要包含 markdown 代码块或其他解释性文字。""",
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

        return result

    except (json.JSONDecodeError, AttributeError, ValueError) as e:
        print(f"JSON 解析失败: {e}")
        return {
            "background_color": "#FFFFFF",
            "elements": [
                {
                    "type": "text",
                    "content": "解析幻灯片失败，请查看后端日志。",
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
            raise HTTPException(status_code=400, detail="未能提取到图片。")

        llm = get_llm()
        target_presentation = Presentation()
        target_presentation.slide_width = Inches(13.333)
        target_presentation.slide_height = Inches(7.5)

        for slide_idx, image_base64 in slide_images:
            analysis = analyze_slide_image(llm, image_base64, slide_idx)
            create_slide_from_analysis(target_presentation, analysis, slide_idx)

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
    return {"status": "ok", "message": "Backend is running."}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
