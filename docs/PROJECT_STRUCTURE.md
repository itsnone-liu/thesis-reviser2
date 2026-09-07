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

### txt 预渲染审计层（2026-09-07，用户拍板：生成→txt后、渲染前独立审计）

- `audit_txt.py`：8 维确定性审计（T1标签变体 graphic/figure/image→渲染器不认即图丢 / T2 LaTeX残留 / T3裸行图注 / T4裸标题+括号参数 / T5孤儿表注 / T6孤儿图注 / T7摘要区标签残留 / T8图表引用悬空），hard 阻断+warn 告警，只报告不改写（改写归 tagguard）
- `renderer.py`：render() 入口已接审计层，hard 直接 ValueError 阻断渲染
- tagguard 扩容：变体标签规范化（graphic/figure/fig/image→drawing，name=→title= 属性映射）
- 验证：50 篇存量 txt 批跑 hard 命中恰好=已修复两案（严涛/张超），零误报

### 漂移治理+空承诺闭环（2026-09-07，两轮架构讨论落地）

- `factcard.py`：事实卡（章节生成时把已确立工程参数注入后续 prompt——"查表"替代"回忆"）+ 章节检查点（章后跑闸门，冲突带反馈重写本章，修复半径=一章，最多重试1次）
- `generator.py`：`_generate_civil`（事实卡+检查点）/`_generate_mechanical`（检查点；事实卡由既有 machine_spec 结构化承担）
- `audit_txt.py` 新增维度：T9 无编号指代解算（如图/图中/下图→同节锚点搜索，无锚=空承诺 hard；类别名词/带编号/引他文排除）、T10 自报清单对账（[FIGURES]块 vs 实际标签）
- `rescue_refs.py`：确定性补偿器（R1 裸标题+括号→drawing 标签；description 只用原文素材不编造；独立于审计层——审计只报告，补偿显式调用）
- `prompts/civil.py`：输出纪律+2（禁止提及不存在的图；章末自报 [FIGURES] 清单）
- `civil_consistency.py`：抽取器副词容差（"均为/约为3.6m"不再漏抽）
- 验证：T9 活性4案例全对+50篇存量零误报；补偿端到端（T4 hard→标签→全绿）；事实卡真实素材抽卡通过；273 docx 全绿
- 两阶段生成（图表清单先行 schema）列为新管线待办；无源审计按用户拍板不做（今后强制保留 txt）
