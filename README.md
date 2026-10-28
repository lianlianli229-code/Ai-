# 🔥 AI 视觉技能互助社区 (AI Visual Skill Community)

基于 **Google Sites / GitHub Pages (前端)** + **Google Sheets & Google Apps Script (后端数据库与 API)** 构建的内部 AI 视觉技能共享、个人技能炉与求助互助平台。

---

## 📁 目录与文件结构

| 文件 | 说明 | 部署位置 |
| :--- | :--- | :--- |
| [`index.html`](./index.html) | **前端完整单页应用代码**（含技能广场、我的技能炉、Google 邮箱身份中心、发布技能、求助广场） | Google Sites 嵌入代码 / GitHub Pages |
| [`Code.gs`](./Code.gs) | **后台 Google Apps Script 脚本代码**（含双向同步、已删除内容清理、多用户技能炉档案、求助互助存储） | Google Sheets → 扩展程序 → Apps Script |

---

## 🚀 部署与更新指南

### 1. 后台部署 (`Code.gs`)
1. 打开绑定的 Google 表格，点击顶部菜单 **扩展程序 (Extensions)** → **Apps Script**。
2. 将 [`Code.gs`](./Code.gs) 中的全部代码复制并覆盖到编辑器中，按 `Ctrl + S` 保存。
3. 点击右上角 **部署 (Deploy)** → **管理部署 (Manage deployments)** → 点击右上角铅笔图标 **编辑** → 版本选择 **新版本 (New version)** → 点击 **部署**。
4. 刷新 Google 表格后，顶部会出现 **🔥 AI技能社区后台** 菜单，包含：
   - **🚀 一键同步表格数据到前端网站**
   - **🗑️ 前端更新已删除内容**（清理表格中已删除的技能并同步清理前端）
   - **📂 初始化/检查【网站投稿、求助与用户档案表】**

### 2. 前端部署 (`index.html`)
- **方式 A（Google Sites）**：打开 Google Sites 编辑页面 → 双击嵌入代码框 → 粘贴 [`index.html`](./index.html) 全文 → 点击右上角 **发布 (Publish)**。
- **方式 B（GitHub Pages）**：直接将本仓库上传至 GitHub，在仓库 `Settings` → `Pages` 中选择 `main` 分支即可自动托管访问。
