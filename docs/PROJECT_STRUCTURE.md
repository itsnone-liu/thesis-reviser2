# 项目结构

## 顶层目录

- `core.py`：基础层，封装 LLM 调用、标签解析、图表/图纸生成、DOCX 排版
- `profile.py`：从题目/对象生成论文画像
- `generator.py`：按章节生成纯文本论文
- `renderer.py`：把纯文本渲染为 DOCX
- `reviser.py`：已有论文的诊断与重写
- `delivery/`：交付层，包含批处理、服务、烟测和巡检
- `prompts/`：各专业的章节提示词
- `mechanical_spec.py` / `mechanical_cad.py`：机械类一致性约束和图纸生成
- `svg_engineering_drawing.py` / `svg_mold_drawing.py`：工程图生成辅助
- `output/`：本地生成结果与样例输出，默认不纳入版本库

## 交付层

- `delivery/service.py`：单个任务线上服务
- `delivery/batch.py`：批处理入口
- `delivery/runner.py`：流程编排
- `delivery/smoke_test.py`：自动巡检
- `delivery/audit.py`：诊断和检查

## 建议的使用顺序

1. `profile.py`
2. `generator.py`
3. `renderer.py`
4. `delivery/smoke_test.py`
5. `reviser.py`

## 目录约定

- 源码放根目录
- 说明文档放根目录或 `docs/`
- 生成物统一进 `output/`
- 测试和发布前检查统一走 `delivery/`
