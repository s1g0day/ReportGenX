# Update Log

## v0.21.5 (2026-06-29)

- **系统设置**：增加城市和区域字段到基本设置面板
- **工具箱**：增加 supplierName UI 设置项、Web 服务开发模式提示、多行定级依据输入
- **模板修复**：修复 `intranet_vuln` / `Attack_Defense` 空数据表格残留空壳表头、含默认下拉值但无实际数据的空行误判、统计计数包含空行导致虚高，以及 `intranet_vuln` 空 section 文本因预采集索引过期插入到标题上方、渗透描述合并到标题和图片描述渲染
- **模板改进**：intranet_vuln / Attack_Defense 数组表格（可控服务器列表、数据库连接信息表、数据统计信息表）不再自动预填默认类型和空行，改为空表手动添加
- **其他**：更新 combined.db 和 single_vuln_report 模板

---

## v0.21.4 (2026-06-08)

- **Bug 修复**：修复生产环境模板删除失败（user_template_ids 路径比较将完整模板路径与父目录做 == 比较，导致用户导入的模板删除时返回 404）
- **AppData 升级迁移**：新增重装场景下 AppData 配置自动升级机制，防止版本跨度大时配置文件不兼容

---

## v0.21.3 (2026-06-08)

- **模板管理**：模板详情弹窗增加「设为默认」按钮（JSON 文件持久化 order，重启不丢失）；导入同名模板静默替换并通知被替换列表
- **系统设置**：增加 Web UI 开关（控制 `/ui` 路由，默认关闭，即时生效，middleware 拦截）
- **Bug 修复**：修复 Windows 批量导入导出路径分隔符不兼容（3 处 `\` → `/` 规范化）；修复默认模板重启失效（引入 `template_order.json` 原子读写）；修复导入 overwrite=false 导致替换永远不触发
- **其他**：`backend/templates/_deleted/` 加入 gitignore

---

## v0.21.2 (2026-06-04)

- **文档同步**：修复 README.md / AGENTS.md 文件列表，更新架构分析和项目概览文档
- **Widget 增强**：vuln_list 通过 `count_levels` 配置驱动漏洞统计（含 `notifyDataChanged` 初始化），`pre_compute` 联动预计算注入变量上下文，`has_widgets` 标记避免模板 404
- **模板修复**：`vuln_url` 改为多行输入，移除废弃字段（`vuln_description`/`vuln_suggestion`/`vuln_reference`），内网渗透报告改进，Pydantic `extra=allow` 兼容旧模板

---

## v0.21.1 (2026-06-03)

- **AppData 迁移**：DB、报告、日志、配置从安装目录迁移到 `%APPDATA%/ReportGenX/`，重装不再丢失数据
- **NSIS 数据保留**：安装/重装/卸载时询问是否保留已有数据，修复 userData 目录名
- **模板系统重构**：合并 4 个 vuln_list.js 为共享 widget（净减少 2227 行），schema 驱动渲染和声明式映射，解除 form-renderer 硬编码
- **模板修复**：修复用户导入模板 handler 路径、`in` 操作符 TypeError（24 处 None 防护）、内网/单漏洞报告改进

---

## v0.21.0 (2026-06-02)

- 重构更新通知 UI：toast + 脉冲点替代横幅和弹窗
- **⚠ 注意**：v0.21.0 之前的版本在更新时会清空数据，更新前请备份

---

## v0.20.2 (2026-06-02)

- 使用 `kill+wait` 释放文件锁，修复更新安装时进程未完全关闭的问题
- 跳过 NSIS 运行中应用检测
- 使用 `process.exit(0)` 实现更新时立即关闭
- 将后端清理移到 `window-all-closed` 事件

---

## v0.20.1 (2026-05-30)

- 添加 `PLUGIN descriptor` 回退逻辑
- 修复 API 响应解析（`resp.templates`）
- CI 修复：Playwright Electron 路径、`--no-sandbox`、electron 二进制下载

---

## v0.18.7 (2026-05-09)

- Windows/macOS 平台前缀优化 artifact 命名
- 移除 Linux 引用
- 升级 `softprops/action-gh-release` 到 v3

---

## v0.17.7 (2026-02-04) — Electron + FastAPI 重构

初次基于 Electron + FastAPI 的完整重构版本：
- 模板驱动报告生成（schema.yaml + handler.py + template.docx）
- 漏洞库、ICP 备案库管理
- 报告合并、数据备份、配置管理
- 支持 Windows / macOS / Linux

---

## v0.12.2 (2025-12-29) — PyQt6 版本

首个公开发布版本，基于 PyQt6。
