# PPTX Rebuilder - Vibe Coding Guidelines

## Project Overview

This project converts image-based presentation files (PPTX) into traditional, editable PPTX files. It uses a React frontend for file uploading/downloading, and a Python backend (FastAPI) to process the conversion using LangChain and LLM.

## Tech Stack

- **Frontend**: React, Vite (Javascript/JSX).
- **Backend**: Python 3.12+, FastAPI, `uv` for package management.
- **AI/Logic**: LangChain (`langchain-anthropic`), `python-pptx` for generating slides, `pdf2image` or similar for extracting images from the source PPTX.

## Implementation Steps for Claude Code

1. **Slide Extraction**: Parse the uploaded PPTX, extract each slide as a high-res image.
2. **Vision Analysis**: Pass each image to LLM Vision via LangChain. Prompt the model to extract:
   - Layout information (bounding boxes for titles, content, images).
   - Text content (including hierarchies/bullet points).
   - Style hints (font sizes, colors if possible).
3. **Reconstruction**: Use `python-pptx` to create a new presentation. For each analyzed slide, recreate text boxes and standard PPTX shapes based on the JSON output from LLM.
4. **Integration**: Connect the FastAPI endpoint `/api/convert` to receive the file, run the pipeline, and return the new PPTX as a downloadable blob.

## Coding Rules

- ALWAYS output complete files when making updates.
- Keep the FastAPI endpoints RESTful and handle exceptions gracefully (return standard HTTP 500 errors with details).
- Frontend must be responsive and handle loading states (conversion takes time).
- Ensure generated frontend code uses standard Simplified Chinese fonts (`font-family: 'PingFang SC', 'Microsoft YaHei', sans-serif;`).
- Backend uses `uv` for dependency management (`uv add <package>`, `uv run <script>`). Do not use `pip` or `poetry`.
- 使用中文回复。
- 代码注释使用中文，解释简洁直接，可以举例说明。
- 前端代码应使用现代 React 语法（函数组件、Hooks），避免使用过时的类组件。
- 后端代码应遵循 PEP 8 风格指南，保持代码清晰易读。
- 在处理文件上传和下载时，确保安全性，避免潜在的文件注入攻击。
- 在前端和后端之间传递数据时，使用 JSON 格式，并确保正确处理编码和解码。
