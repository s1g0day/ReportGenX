# Update Log

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
