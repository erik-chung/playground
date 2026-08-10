# QQ群成员导出工具

使用 **NapCat.Shell**（基于 NTQQ）提供 OneBot v11 HTTP API，Python Flask Web 程序连接 API 实现群成员批量导出。

## 功能特性

- Web 界面操作，无需命令行
- 群列表按创建时间倒序展示
- 显示群头像、名称、成员数、群主头像及昵称
- 超过 5 个群自动分页
- 一键导出群成员名单为 Excel
- 按入群时间正序排列

## 快速开始

### 第一步：下载 NapCat.Shell

1. 访问 https://github.com/NapNeko/NapCatQQ/releases
2. 下载最新版 `NapCat.Shell.zip`
3. 解压到任意目录（如 `D:\NapCat.Shell\`）
4. 运行 `NapCat.Shell.exe`，扫码登录 QQ

### 第二步：配置 HTTP API

1. 浏览器打开 NapCat WebUI：http://127.0.0.1:6099
2. 输入 token 登录（token 在 `config/webui.json` 中查看）
3. 进入 **OneBot11 配置** → **HTTP 服务**
4. 添加 HTTP 服务器：
   - 主机：`0.0.0.0`
   - 端口：`8081`
   - 启用：√
   - 其余默认
5. 保存配置，重启 NapCat.Shell

### 第三步：运行 Python 程序

双击 `启动.bat`，或手动运行：

```bash
pip install flask flask-cors openpyxl requests
python web_exporter.py
```

### 第四步：使用

打开浏览器访问 http://127.0.0.1:8080

---

## 文件说明

- `web_exporter.py` - Python Web 主程序
- `启动.bat` - Windows 快速启动脚本
- `templates/index.html` - Web 界面
- `exports/` - 导出的 Excel 文件存放目录

---

## Excel 导出格式

导出的 Excel 包含以下列，按入群时间正序排列：

| 列名 | 说明 |
|------|------|
| QQ号 | 成员的QQ号码 |
| 昵称 | QQ昵称 |
| 群名片 | 在群内的昵称 |
| 角色 | 群主/管理员/成员 |
| 入群时间 | 加入群的时间 |
| 最后发言 | 最后发言时间 |
