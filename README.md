# PPTX Rebuilder

将图片型 PPT 转换为可编辑文本的 PPT。使用 Claude Opus 4.5 视觉模型分析幻灯片图像，提取文本与布局信息，并重建为标准 PPTX 格式。

## 功能特点

- 上传由图片构成的 PPTX 文件
- 使用 AI 视觉模型进行 OCR 和布局分析
- 自动重建为可编辑的文本型 PPTX
- 支持标题、正文、列表等常见布局

## 技术栈

| 层级     | 技术                                    |
| -------- | --------------------------------------- |
| 前端     | React 18 + Vite                         |
| 后端     | Python 3.12+ + FastAPI                  |
| AI       | SAP Generative AI Hub + Claude Opus 4.5 |
| PPT 处理 | python-pptx                             |

## 项目结构

```
pptx-rebuilder/
├── frontend/          # React 前端
├── backend/           # FastAPI 后端
└── README.md
```

## 快速开始

### 环境要求

- Node.js 18+
- Python 3.12+
- [uv](https://docs.astral.sh/uv/) (Python 包管理器)
- SAP AI Core 服务实例（用于访问 Claude 模型）

### 配置 SAP AI Core 凭证

有两种方式配置 SAP AI Core 凭证：

#### 方式一：配置文件（推荐）

创建 `~/.aicore/config.json` 文件：

```json
{
    "AICORE_CLIENT_ID": "your-client-id",
    "AICORE_CLIENT_SECRET": "your-client-secret",
    "AICORE_AUTH_URL": "https://your-subdomain.authentication.sap.hana.ondemand.com/oauth/token",
    "AICORE_BASE_URL": "https://api.ai.your-region.cfapps.sap.hana.ondemand.com/v2",
    "AICORE_RESOURCE_GROUP": "default"
}
```

#### 方式二：环境变量

```bash
export AICORE_CLIENT_ID="your-client-id"
export AICORE_CLIENT_SECRET="your-client-secret"
export AICORE_AUTH_URL="https://your-subdomain.authentication.sap.hana.ondemand.com/oauth/token"
export AICORE_BASE_URL="https://api.ai.your-region.cfapps.sap.hana.ondemand.com/v2"
export AICORE_RESOURCE_GROUP="default"
```

> 这些凭证可以从 SAP BTP Cockpit 中 AI Core 服务实例的 Service Key 获取。

### 启动后端

```bash
cd backend
uv sync           # 安装依赖
uv run main.py    # 启动服务，默认端口 8000
```

### 启动前端

```bash
cd frontend
npm install       # 安装依赖
npm run dev       # 启动开发服务器，默认端口 5173
```

### 访问应用

打开浏览器访问 http://localhost:5173

## API 接口

### POST /api/convert

上传 PPTX 文件并转换。

**请求**: `multipart/form-data`，字段名 `file`

**响应**: 转换后的 PPTX 文件（二进制流）

**示例**:

```bash
curl -X POST http://localhost:8000/api/convert \
  -F "file=@input.pptx" \
  -o output.pptx
```

### GET /api/health

健康检查接口，同时验证 SAP AI Core 配置状态。

**响应**:

```json
{ "status": "ok", "message": "Backend is running with SAP AI Core configured" }
```

如果 SAP AI Core 未完全配置：

```json
{
  "status": "warning",
  "message": "Backend is running but SAP AI Core not fully configured",
  "missing_env_vars": ["AICORE_CLIENT_ID", "..."]
}
```

## 开发状态

- [x] 前端上传/下载界面
- [x] 后端 API 框架
- [x] SAP Generative AI Hub 集成
- [ ] 幻灯片图像提取
- [ ] AI 视觉分析与布局识别
- [ ] PPT 重建逻辑
