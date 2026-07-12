# thesis-reviser-next

论文自动生成、修改与渲染工具链，面向本科毕业论文的批量生产和单篇交付。

支持四类论文：
- 管理
- 设计
- 机械
- 土木

核心流程：

```text
profile.py -> generator.py -> renderer.py -> DOCX
                    ^
                    |
               reviser.py
```

## 特点

- 纯文本生成和最终渲染分离，便于排查标签和排版问题
- 支持 `<chart/>`、`<table/>`、`<drawing/>` 三类结构化标签
- 支持批量生成、单篇生成、已有论文修改
- 土木、机械、设计类具备图纸渲染链路
- 交付层带有 smoke test 和结构巡检

## 仓库结构

详细说明见 [docs/PROJECT_STRUCTURE.md](docs/PROJECT_STRUCTURE.md)。

## 环境要求

- Python 3.10+
- `python-docx`
- `requests`
- `Pillow`
- `matplotlib`
- `numpy`
- 可选：`libreoffice`、`poppler-utils`
- 可选：`freecad-python3`

## 环境变量

发布版不再在代码中硬编码密钥，需要通过环境变量注入：

- `DEEPSEEK_API_KEY`
- `IMAGE_API_KEY`

示例：

```bash
export DEEPSEEK_API_KEY="your-deepseek-key"
export IMAGE_API_KEY="your-image-key"
```

## 快速开始

### 1. 生成画像

```bash
python profile.py --type 管理 --title "A公司财务管理问题与对策研究" --object "A公司" -o output/A公司_profile.json
```

### 2. 生成论文文本

```bash
python generator.py --profile output/A公司_profile.json --type 管理 -o output/A公司_paper.txt
```

### 3. 渲染 DOCX

```bash
python renderer.py --input output/A公司_paper.txt --output output/A公司_paper.docx --type 管理
```

## 常用命令

### 单篇修改

```bash
python reviser.py --input /path/to/original.docx --output ./output/revised --type 土木
```

### 批量生成

```bash
python batch_generate.py --input output/batch_titles.json --output-dir output/batch_run --type 管理
```

### 土木批量生成

```bash
python batch_generate_civil.py --input output/batch_titles.json --output-dir output/batch_run --type 土木
```

## 质量检查

```bash
python -m delivery.smoke_test
```

更多发布前检查规则见 [docs/RELEASE.md](docs/RELEASE.md)。

## 说明

- `output/` 默认忽略，不纳入版本库
- 仓库中的样例文档仅用于本地验证，不建议作为发布产物提交
- 代码中的模板、标签规范和巡检规则都面向“可交付”场景设计
