# PROJECT_OVERVIEW

> 面向接手开发者的快速上下文文档
> 更新时间：2026-06-04
> 当前版本：0.21.1

## 1. 项目定位

`ReportGenX` 是一个 Electron + FastAPI 的桌面应用，用于生成安全报告与维护配套知识库。

核心能力：

- 模板驱动报告生成（`schema.yaml` + PLUGIN descriptor `handler.py` + `GenerationContext`）
- 漏洞库管理（增删改查、Excel 导入导出）
- ICP 信息管理（增删改查、批量删除）
- 报告列表、删除、合并
- 模板导入/导出/热加载

---

## 2. 技术栈与运行形态

### 前端

- 原生 HTML/CSS/JS（无 React/Vue）
- 入口：`src/index.html`
- 模块：`src/js/*.js`

### 桌面容器

- Electron 主进程：`main.js`
- Preload 桥接：`preload.js`

### 后端

- FastAPI：`backend/api.py`（~2174 行，单文件全量 API）
- 核心逻辑：`backend/core/*`
- SDK 门面：`core/__init__.py`（从 backend.core.* 重导出）
- 插件运行时：`backend/plugin_host/runtime.py`
- 模板目录：`backend/templates/*`（6 个模板）
- 数据库：`backend/data/combined.db`（SQLite）

---

## 3. 目录快速地图（常改文件）

- `main.js`：后端进程生命周期、启动握手、外链安全策略
- `preload.js`：将 `apiBaseUrl`/`appApiToken` 注入渲染层
- `src/js/api.js`：统一 API 调用与 token header 注入
- `src/js/form-renderer.js`：动态表单引擎（报告生成主流程）
- `src/js/toolbox.js`：工具箱业务（系统设置/模板管理/报告合并等）
- `src/js/template-manager.js`：模板管理 UI 逻辑
- `src/js/vuln-manager.js`：漏洞库管理
- `src/js/crud-manager.js`：通用 CRUD 抽象
- `backend/api.py`：API 入口、中间件、路由
- `backend/core/template_manager.py`：模板扫描与依赖检查（含安全审计）
- `backend/core/generation_context.py`：模板服务注入层（~947 行）
- `backend/core/schema_loader.py`：YAML → Pydantic 解析
- `backend/shared-config.json`：server/security/paths/plugin_runtime 共享配置

---

## 4. 启动与通信链路（执行顺序）

1. Electron 启动 `main.js`
2. `main.js` 启动后端进程并注入 `APP_API_TOKEN`
3. 主进程调用 `GET /api/health-auth`（带 `X-App-Token`）验证令牌绑定
4. 验证通过后创建窗口并加载前端
5. `preload.js` 注入 `window.electronConfig` 与 `window.electronAPI`
6. 渲染层通过 `window.AppAPI` 与 FastAPI 通信
7. 生成报告请求进入 `PluginRuntime.execute()` → PLUGIN descriptor → `GenerationContext` → 落盘并返回下载路径

---

## 5. 当前安全/运行时边界

`backend/api.py` 中的 `app_token_middleware` 规则：

- `/api/*` 的 `POST/PUT/PATCH/DELETE` 默认需要 `X-App-Token`
- 受保护 GET：
  - `/api/backup-db`
  - `/api/health-auth`
  - `/api/templates/{template_id}/export`
- `GET /api/health` 不鉴权（基础存活探针）

重要说明：

- 当前 `plugin_runtime` 配置接口为：
  - `GET /api/plugin-runtime-config`（读取）
  - `POST /api/plugin-runtime-config`（更新，受 token 保护）
- 当前没有管理员会话口令门禁（无 `admin-session` / `X-Admin-Session`）

---

## 6. 配置体系

### `backend/config.yaml`

- 业务配置（版本、风险等级、下拉选项等）
- 数据库路径（`vul_or_icp`）

### `backend/shared-config.json`

关键字段：

- `server.host/server.port`
- `security.external_protocols` / `security.external_hosts`
- `paths.open_folder_allowlist`
- `plugin_runtime.*`（11 个配置键：mode、isolated 灰度、fallback、metrics 等）

---

## 7. 模板系统

### 模板目录结构

```
backend/templates/<template_id>/
├── schema.yaml        # 表单定义（14 种字段类型、数据源、行为、验证规则）
├── handler.py         # PLUGIN descriptor + execute() + pure functions
├── template.docx      # Word 模板（含 #placeholder# 占位符）
├── runtime.yaml       # 日志/输出配置
└── widgets/           # 可选：自定义 JS/CSS 组件
```

模板级 Widget 位于模板的 `widgets/` 子目录，项目级共享 Widget 位于 `backend/widgets/`（当前包含通用 `vuln_list.js` 漏洞列表组件，合并了 4 个模板的近似代码）。

### 架构模式

当前标准：**PLUGIN descriptor + GenerationContext**（不再使用 `BaseTemplateHandler` 类继承）。

```python
# handler.py
PLUGIN = {
    "id": "template_id",
    "execute": execute,   # execute(data, output_dir, template_manager, config, template_id) -> dict
}

def generate(data, ctx):  # ctx = GenerationContext — 提供所有框架服务
    doc = ctx.load_document()
    ctx.replace_text(replacements)
    ctx.process_single_image('#ph#', data.get('img'))
    return True, ctx.save(filename), "OK"
```

### 当前模板（6 个）

| 模板 ID | 名称 | 说明 |
|---------|------|------|
| `vuln_report` | 漏洞报告 | 标准漏洞报告，含证据截图 |
| `intrusion_report` | 入侵痕迹报告 | 入侵痕迹报告，含时间线 |
| `penetration_test` | 渗透测试报告 | 完整渗透测试报告，含漏洞表/风险图/目录 |
| `Attack_Defense` | 攻防演练报告 | 攻防演练，含服务器类型/DB 连接/数据统计 |
| `single_vuln_report` | 单个漏洞报告 | 最简模板，从漏洞库快速选取填充 |
| `intranet_vuln` | 内网测试报告 | 内网漏洞测试报告，含渗透入口、互联网/内网渗透路径、数据汇总 |

---

## 8. 验证与 CI

- Electron 冒烟：`npm run test:e2e:smoke`（Playwright）
- CI 冒烟工作流：`.github/workflows/ci-smoke.yml`
- 后端单元测试：`npm run test`（已移除，当前为空操作）

---

## 9. 接手建议（最短路径）

建议阅读顺序：

1. `README.md`
2. `main.js`、`preload.js`
3. `src/js/api.js`
4. `src/js/form-renderer.js`
5. `backend/api.py`
6. `backend/core/generation_context.py`
7. 任一模板目录（推荐 `backend/templates/single_vuln_report/` — 最简模板）

高频改动入口：

- 新增字段/表单：改对应模板 `schema.yaml`
- 调整报告生成逻辑：改模板 `handler.py`
- 调整接口行为：改 `backend/api.py`
- 调整工具箱交互：改 `src/js/toolbox.js`
- 运行时配置：`backend/shared-config.json` → `plugin_runtime` 块
