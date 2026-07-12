# Delivery Layer

这个目录是从现有仓库复制出来的交付层，不改原始脚本，专门提供三种能力：

1. 批量直接运行
2. 单任务线上服务
3. 问题梳理与抽检报告

## 批量模式

```bash
python -m delivery.batch --workflow standard --input output/batch_titles.json --output-dir output/batch_run --type 管理
```

土木批处理：

```bash
python -m delivery.batch --workflow civil --civil-args --single 张三
```

标准批处理同样支持土木：

```bash
python -m delivery.batch --workflow standard --input output/batch_titles.json --output-dir output/batch_run --type 土木
```

## 服务模式

```bash
python -m delivery.service
```

服务接口支持 `管理`、`设计`、`机械`、`土木` 四类请求。

默认监听 `0.0.0.0:8000`，可用环境变量调整：

- `THESIS_REVISER_HOST`
- `THESIS_REVISER_PORT`

## 迁移说明

旧的三类入口仍兼容，但现在土木已经独立成第四类。对外调用时请直接传 `土木`，不要再让土木题目落回 `管理` 分支。

## 抽检模式

```bash
python -m delivery.audit
```

默认扫描 `修改记录.md` 和 `今日进度_20260704.md`，也可以额外传入目录或文件。
