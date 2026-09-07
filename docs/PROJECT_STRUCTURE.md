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

## 终版 docx-only 审计与修复（2026-09-06/07）

- `audit_docx_only2.py`：论文终版 273 篇无源审计（缺图/引用悬空/封面/目录/标题编号）。要点：body 内嵌 `w:drawing`/`w:pict` 计数而非 media 数；命名空间前缀归一（Word/WPS 另存后的 `ns0:` 文档）；引用识别三防——构词防（简图/插图/附图/图纸/大样图/示意图/流程图/路线图/效果图+数字是名词）、TOC 粘连防（目录行"标题+页码"非引用）、跨词拼接防（编号内禁空白）
- `fix_final_images.py`：docx 级四类手术（A 图注在图缺→AI 生图插图注前；B 图注图都缺→生图+造注插引用句后；C 图注补"图章-序"编号，双侧探测方位；D 引用重编号）。幂等注意：C 类重复执行会把已编号图注再编号
- 配套报告：`/root/project/workspace/{终版审计_终验0907,加固审计_0907,图片修复报告_0907}.csv`；修订输出 `/root/project/workspace/论文终版_修订0907/`
