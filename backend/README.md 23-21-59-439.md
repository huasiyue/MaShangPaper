# Backend

`backend/` 是 MaShangPaper 的 FastAPI 服务端，负责：

- 接收 Markdown 内容并生成 Word 初稿
- 接收 Word 文档并生成格式审查报告
- 对 Word 文档套用学校模板并导出结果包
- 对前端暴露统一 HTTP API

当前默认模板仅支持扬州大学。

## 运行环境

后端默认使用 conda `pytorch` 环境：

```bash
/opt/anaconda3/envs/pytorch/bin/python
```

安装依赖：

```bash
cd /Users/yun/Documents/project/trea_project/github/MaShangPaper/backend
/opt/anaconda3/envs/pytorch/bin/python -m pip install -r requirements.txt
```

启动服务：

```bash
cd /Users/yun/Documents/project/trea_project/github/MaShangPaper/backend
/opt/anaconda3/envs/pytorch/bin/python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

## 目录结构

```text
backend/
├── app/
│   ├── api/                  # FastAPI 路由层
│   ├── config/               # 学校配置
│   ├── core/                 # 基础设置
│   ├── integrations/         # legacy 脚本接入层
│   ├── schemas/              # Pydantic 传输模型
│   ├── services/             # 服务层
│   │   └── formatters/       # 学校格式器
│   └── main.py               # 应用入口
├── requirements.txt
└── temp/                     # 临时文件目录
```

## 代码职责

### `app/main.py`

负责：

- 创建 FastAPI 应用
- 配置 CORS
- 注册异常处理器
- 挂载路由

### `app/api/documents.py`

负责：

- 定义 `/api/documents/*` 接口
- 解析表单和文件上传
- 做基础参数校验
- 调用服务层并返回文件流或 JSON

### `app/services/document_pipeline.py`

负责：

- 组织 `convert`、`review`、`format` 三条主流程
- 管理上传文件落盘
- 管理临时输出路径
- 生成格式化结果 ZIP

### `app/integrations/legacy_scripts.py`

负责：

- 动态加载 `docs/data/convert_to_word.py`
- 动态加载 `docs/data/yzu_thesis_formatter.py`
- 作为当前阶段的 legacy 兼容层

### `app/services/formatters/yzu.py`

负责：

- 调用扬州大学格式化脚本
- 输出结构化审查结果

### `app/schemas/documents.py`

负责：

- 定义接口响应模型
- 定义审查问题结构
- 统一错误返回结构

## API 接口

### `GET /`

用途：

- 根健康检查

返回：

```json
{
  "status": "ok",
  "school_support": ["yzu"]
}
```

### `GET /api/documents/health`

用途：

- 文档服务健康检查

返回：

```json
{
  "status": "ok",
  "school_support": ["yzu"]
}
```

### `POST /api/documents/convert`

用途：

- 将 Markdown 内容转换为 Word 初稿

表单参数：

- `content`: Markdown 文本
- `school_id`: 当前仅支持 `yzu`
- `thesis_type`: `thesis` 或 `design_report`

返回：

- `.docx` 文件流

### `POST /api/documents/review`

用途：

- 对上传的 Word 文档做格式审查

表单参数：

- `file`: `.doc` 或 `.docx`
- `school_id`: 当前仅支持 `yzu`
- `thesis_type`: `thesis` 或 `design_report`

返回：

```json
{
  "filename": "sample.docx",
  "school_id": "yzu",
  "thesis_type": "thesis",
  "total_issues": 6,
  "error_count": 0,
  "warning_count": 1,
  "info_count": 5,
  "issues": [],
  "report_text": "..."
}
```

### `POST /api/documents/format`

用途：

- 将上传的 Word 文档按学校模板格式化

表单参数：

- `file`: `.doc` 或 `.docx`
- `school_id`: 当前仅支持 `yzu`
- `thesis_type`: `thesis` 或 `design_report`

返回：

- ZIP 文件流

ZIP 内包含：

- 格式化后的 `.docx`
- `review.json`
- `review.txt`

## 当前限制

当前版本有以下限制：

- 仅支持 `yzu`
- 格式器仍通过 `docs/data` 中的 legacy 脚本接入
- `convert` 工作流当前更适合“论文 Markdown 子集”，不是通用 Markdown 渲染器
- 暂未加入自动化单元测试

## 开发建议

继续扩展后端时，建议优先保持以下原则：

- 路由层只处理 HTTP 相关逻辑
- 服务层负责工作流编排
- 学校差异尽量放在格式器和配置层
- 输入输出尽量通过 Pydantic schema 明确约束
- 新学校优先新增配置和 formatter，而不是复制整条流程

