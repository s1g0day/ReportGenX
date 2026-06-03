# 部署与运维指南

> 更新日期：2026-06-04
> 适用版本：0.21.1

## 1. 部署模式

ReportGenX 当前推荐 **Electron 单机模式**（桌面应用 + 本地 FastAPI）。

- 目标场景：单机离线/内网终端使用
- 运行链路：Electron 主进程拉起本地后端，渲染层通过 `http://127.0.0.1:<port>` 调用 API

> 说明：文档中的"Web 服务化"只作为扩展方案参考。当前默认安全链路依赖 Electron 注入 `X-App-Token`，不适合作为通用开放 Web 服务直接暴露。

---

## 2. 环境要求

### 2.1 运行要求（桌面端）

- CPU：2 核及以上
- 内存：4GB 及以上
- 磁盘：至少 500MB 可用空间（报告与日志）

### 2.2 开发要求

- Node.js >= 16
- Python >= 3.9（CI 使用 3.10）

---

## 3. 本地开发运行

1. 安装前端依赖

```bash
npm install
```

2. 安装后端依赖

```bash
cd backend
pip install -r requirements.txt
```

3. 启动后端（可选，Electron 启动时也会尝试拉起）

```bash
cd backend
uvicorn api:app --host 127.0.0.1 --port 8000 --reload
```

4. 启动 Electron

```bash
npm run start
```

`npm run start` 会自动执行 `prestart` 钩子同步版本号到 `shared-config.json`。

---

## 4. 打包发布

### 4.1 构建 Python 后端可执行

```bash
pyinstaller --noconfirm backend/api.spec
```

输出目录：`backend/dist/`

### 4.2 构建 Electron 安装包

```bash
npm run dist
```

构建前会自动执行 `predist` 钩子（同步版本 + 校验版本一致性）。

常用目标：

- Windows x64：`npm run dist -- --win --x64`
- macOS ARM：`npm run dist -- --mac --arm64`
- macOS x64：`npm run dist -- --mac --x64`

输出目录：`dist/`（安装包 `.exe` / `.dmg` / `.zip`）

### 4.3 版本同步

发布前务必校验版本一致性：

```bash
npm run check-version
```

版本号来源：`package.json` → `scripts/sync-version.js` 同步到 `backend/shared-config.json`。

### 4.4 构建产物结构

打包后的应用目录：

```
ReportGenX/
├── ReportGenX.exe          # Electron 可执行
├── resources/
│   ├── api.exe             # PyInstaller 打包的后端
│   ├── config.yaml         # 业务配置
│   ├── shared-config.json  # 运行时共享配置
│   ├── templates/          # 模板目录
│   ├── widgets/            # 共享 Widget 目录
│   └── data/               # SQLite 数据库
└── ...
```

### 4.5 extraResources 配置

打包时通过 `extraResources` 将运行时必需目录复制到应用资源中：

```json
// package.json → build.extraResources
[
  { "from": "backend/dist",      "to": "backend/dist" },
  { "from": "backend/config.yaml", "to": "backend/config.yaml" },
  { "from": "backend/shared-config.json", "to": "backend/shared-config.json" },
  { "from": "backend/templates", "to": "backend/templates" },
  { "from": "backend/widgets",   "to": "backend/widgets" },
  { "from": "backend/data",      "to": "backend/data" }
]
```

- `backend/templates`、`backend/widgets`、`backend/data` 在打包时随应用发布
- 首次运行后，应用将 `data`、`templates`、`widgets` 复制到系统 AppData 用户目录，后续优先读取 AppData 中的副本，支持运行时自定义和数据持久化

### 4.6 NSIS 安装器数据保留

Windows 安装器配置中使用 `deleteAppDataOnUninstall: false` 保留卸载时的用户数据：

```json
// package.json → build.nsis
{
  "deleteAppDataOnUninstall": false,
  "include": "build/installer.nsh"
}
```

卸载应用时不会删除 `AppData` 中的数据库、模板等用户数据，重新安装后自动恢复。

---

## 5. CI 发布流程

仓库通过 GitHub Actions 自动化发布（`.github/workflows/release.yml`）：

1. **触发**：推送 `v*.*.*` 标签或手动触发
2. **前置检查**：`check_tag_version` — 验证 tag 版本与代码版本一致
3. **构建矩阵**：
   - `win-x64`：打包 `api.exe` (PyInstaller) + Electron NSIS 安装包
   - `mac-x64` / `mac-arm64`：打包 `api` + Electron DMG/ZIP
4. **上传**：构建产物自动发布到 GitHub Release

可选 Windows 代码签名：设置 `CSC_LINK` / `CSC_KEY_PASSWORD` secrets。

CI 冒烟（`.github/workflows/ci-smoke.yml`）在 push 和 PR 时自动运行后端测试 + Electron 冒烟检查。

---

## 6. 运行时安全边界（必须知晓）

### 6.1 Token 策略

后端 `app_token_middleware` 当前规则：

- `/api/*` 的 `POST/PUT/PATCH/DELETE` 默认要求 `X-App-Token`
- 受保护 GET：
  - `/api/backup-db`
  - `/api/health-auth`
  - `/api/templates/{template_id}/export`
- `GET /api/health` 不鉴权（基础存活探针）

### 6.2 启动握手

Electron 主进程在开窗前调用 `GET /api/health-auth`（带 `X-App-Token`）：

- 返回 200：继续启动
- 返回 403：视为"端口已有旧后端实例且 token 不匹配"，直接退出并提示

### 6.3 runtime 配置接口

- `GET /api/plugin-runtime-config`：读取
- `POST /api/plugin-runtime-config`：更新（token 保护）

### 6.4 路径安全

报告输出仅在 `backend/output/report/` 下，受规范化校验。`open-folder` API 仅允许配置白名单中的路径（`output/report`, `output/temp`, `output`）。

### 6.5 外链安全

`shared-config.json` 中 `security.external_hosts` 白名单控制允许打开的 URL：
- `github.com`, `www.github.com`
- 仅允许 `https:` 协议

---

## 7. 运维操作

### 7.1 模板更新

1. 将模板目录放入 `backend/templates/`
2. 调用 `POST /api/templates/reload` 热加载
3. 若新增了模板自定义路由（`router = APIRouter()`），需重启应用后生效

### 7.2 数据库维护

备份建议定期备份以下内容：

- `backend/data/combined.db` — 主数据库（漏洞库、ICP 等）
- `backend/config.yaml` — 业务配置
- `backend/shared-config.json` — 运行时配置
- `backend/templates/` — 模板目录

数据库备份 API：

```bash
curl http://127.0.0.1:8000/api/backup-db -H "X-App-Token: <TOKEN>" -o backup.db
```

### 7.3 常见故障

- `Invalid application token`：通常是旧后端残留；完全退出应用后重启
- 启动即失败且提示 token mismatch：结束占用端口的旧后端进程后重启
- runtime 配置保存 403：检查渲染层是否注入 `window.electronConfig.appApiToken`
- 生成报告下载 404：检查 `backend/output/report/` 目录是否存在且可写
- 打包后模板加载失败：确认 `extraResources` 配置正确打包了 `templates/` 目录

---

## 8. 发布前检查清单

- [ ] `npm run check-version` 通过
- [ ] `npm run test:e2e:smoke` 通过
- [ ] 手动验证：应用启动 → 模板选择 → 表单填写 → 生成报告 → 下载打开
- [ ] 手动验证：工具箱 → 报告合并 → 正常列出报告
- [ ] 手动验证：工具箱 → 系统设置 → runtime 配置读写正常

---

## 9. 参考文档

- [RUNTIME_OPERATIONS_RUNBOOK.md](./RUNTIME_OPERATIONS_RUNBOOK.md) — Runtime 配置与故障排查
- [ARCHITECTURE_LAYER_FLOW_ANALYSIS.md](./ARCHITECTURE_LAYER_FLOW_ANALYSIS.md) — 架构分层分析
- [TEMPLATE_DEV_GUIDE.md](./TEMPLATE_DEV_GUIDE.md) — 模板开发完整指南
