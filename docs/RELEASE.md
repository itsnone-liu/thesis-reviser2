# 发布与交付

## 发布前检查

1. 确认 `python -m delivery.smoke_test` 通过
2. 确认 `git status` 为空
3. 确认 `output/`、缓存、临时文件未被提交
4. 确认代码中没有硬编码密钥

## 构建目标

- 单篇生成：画像 -> 文本 -> DOCX
- 批量生成：名单 -> 批量文本 -> 批量 DOCX
- 线上服务：单任务调用，返回可交付结果

## 推荐发布方式

- GitHub 仓库只放源码、文档、脚本和测试
- 生成结果、样例 DOCX、PDF、图片统一留在 `output/`
- 敏感配置通过环境变量注入

## Git 标签建议

- `v0.1.0`：仓库整理完成
- `v0.2.0`：交付层和巡检稳定
- `v1.0.0`：可对外发布

## 常用检查命令

```bash
git status --short
python -m delivery.smoke_test
```

## 上传前清单

- `README.md` 已更新
- `docs/` 已补齐
- `core.py` 不含硬编码密钥
- `output/` 未被提交
