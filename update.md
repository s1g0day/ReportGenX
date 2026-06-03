# Update Log

## v0.21.1 (2026-06-03)

### 数据持久化
- **用户数据迁移到 OS 用户数据目录**：DB、报告、日志、配置从安装目录迁移到 `%APPDATA%/ReportGenX/`，重装不再丢失数据
- 首次启动自动从安装包复制种子数据到 AppData

### 安装器增强
- **NSIS 数据保留提示**：安装/重装/卸载时询问是否保留已有数据
- 修复 userData 目录名（`report-electron-app` → `ReportGenX`）

### 模板系统
- **修复用户导入模板无法使用**：handler 改用 `get_template_dir()` 双路径解析
- **vuln_save 声明式映射**：模板在 schema.yaml 中声明字段→漏洞库映射
- **合并 4 个 vuln_list.js 为共享 widget**：净减少 2227 行，schema `columns` 驱动渲染
- **架构解耦**：form-renderer 中模板专属硬编码改为 schema 驱动

### 模板改进
- **内网测试报告**：新增报告总结自动生成、修复内网资产渗透路径插入失败
- **单个漏洞报告**：新增网站域名字段、调整字段顺序
- 修复 `in` 操作符 TypeError（24 处 `para.text/cell.text` None 防护）

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
