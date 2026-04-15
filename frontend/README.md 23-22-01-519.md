# Frontend

`frontend/` 是 MaShangPaper 的前端工作台，基于 Vue 3 + Arco Design Vue 实现。

当前页面目标很直接：

- 编辑论文 Markdown
- 实时预览 Markdown
- 调用后端导出 Word 初稿
- 上传 Word 文件并查看格式审查结果
- 下载格式化后的结果包

## 技术栈

- Vue 3
- Vue Router
- Pinia
- Arco Design Vue
- Axios
- Marked
- Vite
- TypeScript

## 安装与启动

安装依赖：

```bash
cd /Users/yun/Documents/project/trea_project/github/MaShangPaper/frontend
npm install
```

开发模式：

```bash
npm run dev
```

生产构建：

```bash
npm run build
```

默认前端地址：

- `http://127.0.0.1:5173`

默认后端地址：

- `http://127.0.0.1:8000`

如果后端地址变化，可以通过 `VITE_API_BASE_URL` 配置。

## 目录结构

```text
frontend/
├── src/
│   ├── api/                      # HTTP 请求封装
│   ├── components/
│   │   ├── editor/               # 编辑区相关组件
│   │   └── review/               # 审查结果组件
│   ├── router/                   # 路由
│   ├── stores/                   # Pinia 状态
│   ├── styles/                   # 全局样式与主题变量
│   ├── views/
│   │   └── editor/               # 页面级组件
│   ├── App.vue
│   └── main.ts
├── index.html
├── package.json
└── vite.config.ts
```

## 页面结构

当前只有一个主页面：

- `src/views/editor/EditorPage.vue`

页面主要分成四块：

1. 项目头图与说明
2. 学校模板与文档类型选择
3. Markdown 编辑区 + 预览区
4. Word 审查与格式化结果区

## 组件说明

### `src/components/editor/SchoolSelector.vue`

负责：

- 选择学校模板
- 选择文档类型

当前支持：

- 学校：扬州大学
- 类型：毕业论文、毕业设计报告

### `src/components/editor/ExportPanel.vue`

负责：

- “导出 Markdown 初稿”
- “审查 Word 格式”
- “格式化并下载 ZIP”

同时展示当前选择的 Word 文件名。

### `src/components/review/ReviewResultTable.vue`

负责：

- 展示审查问题列表
- 区分错误、警告、提示
- 展示文本审查报告

## 状态管理

当前使用 Pinia 管理最小状态集合：

- Markdown 内容
- 学校模板
- 文档类型
- 审查结果

对应文件：

- `src/stores/document.ts`

## API 封装

接口请求统一放在：

- `src/api/http.ts`
- `src/api/documents.ts`

当前封装了三个业务接口：

- `convertMarkdown`
- `reviewWordDocument`
- `formatWordDocument`

## 交互流程

### 导出初稿

1. 在 Markdown 编辑区输入内容
2. 点击“导出 Markdown 初稿”
3. 前端调用 `/api/documents/convert`
4. 浏览器下载 `.docx`

### 审查 Word

1. 选择本地 Word 文件
2. 点击“审查 Word 格式”
3. 前端调用 `/api/documents/review`
4. 页面渲染结构化问题与文本报告

### 格式化 Word

1. 选择本地 Word 文件
2. 点击“格式化并下载 ZIP”
3. 前端调用 `/api/documents/format`
4. 浏览器下载 ZIP 结果包

## 当前限制

当前版本有这些特点：

- 只有一个主页面
- 学校模板先固定为 `yzu`
- Markdown 编辑器目前用 `a-textarea` + `marked` 预览实现，尚未接入更完整的富 Markdown 编辑组件
- 页面已经可构建，但仍属于第一阶段原型

## 后续建议

如果继续演进前端，建议优先做这些事：

1. 接入更强的 Markdown 编辑器组件
2. 增加示例模板和论文结构引导
3. 把审查结果增加筛选、分组和导出能力
4. 为多学校模板准备更清晰的切换入口
