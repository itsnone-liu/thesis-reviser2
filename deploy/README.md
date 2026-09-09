# 汤圆部署说明

1. 先备份老目录与旧输出，不迁移旧生成结果。
2. 在 `/opt/thesis-reviser2` 拉取 `main`，并同步 `case_library/`、模板和 `.env`。
3. 安装 `requirements.txt`。
4. 复制 `thesis-web.service` 到 `/etc/systemd/system/`，执行 `systemctl daemon-reload && systemctl enable --now thesis-web`。
5. 验证 `curl http://127.0.0.1:8000/health`、开放注册、管理员登录、法学出题。

`.env` 只放服务器，不提交 git，权限设为 600。公网直接使用 `http://199.68.217.229:8000`；如后续复用 80 端口，先确认 nginx 当前监听服务，避免影响 OpenClaw。
