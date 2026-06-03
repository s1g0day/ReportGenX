# ReportGenX (Electron + FastAPI)

一个用于生成安全报告的桌面应用，采用 **Electron 前端壳 + 本地 FastAPI 后端 + 模板插件系统**。

> 当前归档版本：`0.21.0`

## 核心能力

- 模板驱动报告生成（`schema.yaml + handler.py + template.docx`）
- 漏洞库管理（CRUD）
- ICP 信息管理（CRUD + 批量删除）
- 报告列表、删除、合并
- 模板导入/导出/热加载
- 工具箱运行时设置（`plugin_runtime`）

---

## 架构分层

- `main.js`：Electron 主进程，负责后端生命周期与启动握手
- `preload.js`：向渲染层注入 `electronConfig`、`electronAPI`
- `src/js/*`：前端业务模块（无框架）
- `backend/api.py`：FastAPI 入口与中间件
- `backend/core/*`：后端核心实现层（业务单一实现来源）
- `backend/plugin_host/runtime.py`：模板执行编排（descriptor/hybrid/legacy/isolated）

---

## 目录结构

```text
├─ backend/
│  ├─ api.py
│  ├─ config.yaml
│  ├─ shared-config.json
│  ├─ core/
│  ├─ plugin_host/
│  ├─ templates/
│  ├─ tests/
│  └─ data/
├─ src/
│  ├─ index.html
│  ├─ styles.css
│  └─ js/
│     ├─ api.js
│     ├─ main.js
│     ├─ form-renderer.js
│     ├─ form-renderer-fields.js
│     ├─ form-renderer-images.js
│     ├─ toolbox.js
│     ├─ template-manager.js
│     ├─ vuln-manager.js
│     └─ crud-manager.js
├─ docs/
│  ├─ ARCHITECTURE_LAYER_FLOW_ANALYSIS.md
│  ├─ PROJECT_OVERVIEW.md
│  ├─ DEPLOYMENT_GUIDE.md
│  ├─ RUNTIME_OPERATIONS_RUNBOOK.md
│  ├─ TEMPLATE_DEV_GUIDE.md
│  └─ TEMPLATE_QUICK_START.md
├─ scripts/
│  ├─ sync-version.js
│  ├─ check-version-sync.js
│  └─ e2e-smoke.js
├─ main.js
├─ preload.js
└─ package.json
```

---

## 运行与开发

### 环境要求

- Node.js >= 16
- Python >= 3.9

### 安装依赖

```bash
npm install
cd backend && pip install -r requirements.txt
```

### 开发模式启动

后端（可选手动）：

```bash
cd backend
uvicorn api:app --host 127.0.0.1 --port 8000 --reload
```

Electron：

```bash
npm run start
```

---

## 安全与启动握手

应用每次启动都会生成随机 `APP_API_TOKEN`，通过 preload 注入渲染层并由后端中间件校验。

后端规则（`backend/api.py`）：

- `/api/*` 的 `POST/PUT/PATCH/DELETE` 默认需要 `X-App-Token`
- 受保护 GET：`/api/backup-db`、`/api/health-auth`、`/api/templates/{id}/export`
- `GET /api/health` 为非鉴权存活探针

主进程在开窗前调用 `GET /api/health-auth`（带 token）确认令牌绑定；若不匹配将直接退出并提示。

---

## 测试与 CI

本地测试：

```bash
npm run test
npm run test:e2e:smoke
```

CI 冒烟：

- 工作流：`.github/workflows/ci-smoke.yml`
- 内容：后端单元测试 + Electron 冒烟

---

## 打包发布

```bash
pyinstaller --noconfirm backend/api.spec
npm run dist
```

发布前版本一致性检查：

```bash
npm run check-version
```

常用目标：

- `npm run dist -- --win --x64`
- `npm run dist -- --mac --arm64`

---

## 文档导航

- 架构分析：`docs/ARCHITECTURE_LAYER_FLOW_ANALYSIS.md`
- 项目概览：`docs/PROJECT_OVERVIEW.md`
- 部署指南：`docs/DEPLOYMENT_GUIDE.md`
- 运行手册：`docs/RUNTIME_OPERATIONS_RUNBOOK.md`
- 模板开发：`docs/TEMPLATE_DEV_GUIDE.md`
- 模板快速入门：`docs/TEMPLATE_QUICK_START.md`
