---
name: grad-design
description: 咸阳职业技术学院毕业设计全流程助手。Use when the user works on a 毕业设计 (graduation design/project), including 毕业设计选题、任务书、设计说明书/正文写作、格式规范排版、指导记录、评分自评、参考文献与图表规范，或提到"电子信息大类/文化艺术大类毕业设计指南""毕业设计规范格式""评分表"。Use ONLY when the task is a graduation design (毕业设计) rather than an ordinary course essay, academic paper, or 毕业论文 with different institutional rules.
---

# 毕业设计助手 (grad-design)

本 skill 依据咸阳职业技术学院《各专业大类毕业设计指南》与《毕业设计规范化要求文件》提供毕业设计全流程指导，覆盖**电子信息大类**（产品/方案/工艺设计类）与**文化艺术大类**（作品/方案设计类）两大专业大类。

## 第一步：永远先澄清（阶段①）

在给出任何选题或写作建议前，必须先确认以下 3 项（缺一不可，已知项可跳过）：

1. **专业大类**：电子信息大类（计算机应用、软件技术、大数据、云计算等）或文化艺术大类（数字媒体艺术、环境艺术等）。
   - 若用户只报了专业名而未报大类：按 `references/topics.md` 的适用专业对照表判断；无法判断时直接询问。
2. **选题类别**：电子信息大类 → 产品设计类 / 方案设计类 / 工艺设计类；文化艺术大类 → 作品设计类 / 方案设计类。
3. **毕业设计形式**：方案 / 物化产品 / 文化作品 / 软件 / 其他（任务书中勾选项）。

## 六阶段工作流

| 阶段 | 目标 | 必读参考 | 可用工具 |
|---|---|---|---|
| ① 澄清 | 确定大类、选题类别、设计形式 | `references/topics.md` | — |
| ② 选题 | 生成候选选题，校验"一人一题"、大小适中 | `references/topics.md` | — |
| ③ 任务书 | 写选题目的意义、任务要求、四周计划安排 | `references/forms.md` | `scripts/generate.py task` |
| ④ 写正文 | 按评分表现状设计结构写设计说明书 | `references/deliverables.md` + `references/rubric.md` | `scripts/generate.py outline` |
| ⑤ 格式自检 | 对照排版规范逐项检查并修复 | `references/format.md` | `scripts/check_format.py` |
| ⑥ 自评归档 | 按评分表打分、给指导记录填写建议 | `references/rubric.md` + `references/forms.md` | — |

每完成一个阶段再进入下一阶段；用户从中间切入时，先定位其所在阶段，再只读该阶段所需参考文件。

## 参考文件索引（按需加载，不要一次性全读）

| 文件 | 何时读 |
|---|---|
| `references/topics.md` | 阶段①②：选题类别、选题要求、选题示例、专业对照 |
| `references/deliverables.md` | 阶段④：各类别成果表现形式与成果要求（正文应包含哪些内容） |
| `references/rubric.md` | 阶段④⑥：5 张评价指标表（科学性/规范性/完整性/实用性）、过程30%+答辩70%、等级线 |
| `references/format.md` | 阶段⑤：字体字号、行距缩进、页码、参考文献、图表、页面的全部排版规则 |
| `references/forms.md` | 阶段③⑥：任务书、指导记录、评分表的字段结构与填写说明 |

## 核心写作原则

1. **评分表反向驱动正文**：写正文前先读 `rubric.md`，把"科学性30 / 规范性20 / 完整性30 / 实用性20"的指标内涵拆成写作清单；正文结构必须覆盖指标要求的全部过程（需求分析→方案分析拟定→技术参数确定→过程记录→效果分析）。
2. **一人一题、任务到人**：选题原则上一人一题；小组课题最多 3 人，任务书中必须写清每个学生的具体分工。
3. **真题真做**：优先来源于企业真实生产/工程/实际项目的课题，大小适中、难易适度、工作量饱满。
4. **格式是硬指标**：规范性占 20 分且评分表单独考察，正文成稿后**必须**跑 `check_format.py` 至 0 错误。

## 脚本用法

所有脚本基于 python-docx，Python 路径为 `python`（Windows）。脚本相对本 skill 目录定位，执行时用绝对路径调用。

```bash
# 格式校验：对正文 .docx 逐项检查报告，返回码非 0 表示存在违规
# （自动识别任务书/指导记录/评分表等表单文档，表单只检查页面设置）
python <skill-dir>/scripts/check_format.py "正文.docx"

# 生成任务书：读取 JSON 输入，填充任务书模板
python <skill-dir>/scripts/generate.py task --input 任务书.json --output 任务书-已填.docx

# 生成正文骨架：按大类+选题类别生成含章节标题与撰写提示的 .docx
python <skill-dir>/scripts/generate.py outline --major electronic --category product --title "XX设计" --output 正文骨架.docx
```

`--major` 取值 `electronic`（电子信息大类）/ `arts`（文化艺术大类）；`--category` 取值 `product`/`scheme`/`process`（电子信息）或 `work`/`scheme`（文化艺术）；`basic` 表示通用骨架（引言/正文/结论）。`outline` 可选 `--student-name --student-id --program --class-name --advisor` 填充封面。

`task` 的 JSON 输入字段（UTF-8）：

```json
{
  "college": "电子信息学院", "program": "电子信息工程技术",
  "class_name": "电信2301", "student_id": "23070101", "advisor": "李老师",
  "title": "题目", "start_date": "2026年9月1日", "end_date": "2026年12月20日",
  "form": "物化产品",
  "purpose": "选题目的意义及主要任务要求（可多段，用 \n 分隔）",
  "weeks": ["第一周……", "第二周……", "第三周……", "第四周……"]
}
```

`form` 限 `方案`/`物化产品`/`文化作品`/`软件`/`其他`；`weeks` 按周数截取（模板固定四周）。

## 模板文件

`templates/` 内含任务书、指导记录、评分表（源自学院原表）与 `正文模板.docx`（引言/正文/结论通用骨架，按规范排版，已通过格式自检）。用户需要电子版表格时优先从这里复制，不要自行发明表格结构；需要按选题类别的章节骨架时用 `generate.py outline` 生成。

## 边界

- 只处理咸阳职业技术学院体系下的毕业设计要求；其他学校/学位论文规范不适用本 skill，应明确告知用户。
- 严禁替学生编造实验数据、虚构企业项目或伪造指导记录签名；内容性建议只做结构与表达层面的指导。
- 原始 `.doc` 参考文件只读，不得修改；一切产出写入用户指定的输出路径。
