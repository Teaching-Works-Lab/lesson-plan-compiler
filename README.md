# Lesson Plan Compiler（教案编制器）

从课程材料、课程大纲与可选的培养方案依据，编制一门课程的标准教案，输出**两份产物**：

- **`教案-working.md`**：可编辑 Markdown，供 AI 与人逐字修改内容；
- **`教案.docx`**：规范 Word，作为最终标准文档（封面 + 课程基本信息 + N 个教学单元表）。

这是一个可复用的 Codex Skill，可单独使用，也可作为 [Teaching Works Lab](https://github.com/Teaching-Works-Lab) 课程工作流的下游（消费 `course-foundation-builder` 产出的课程基座或 `intelligent-manufacturing-syllabus` 的课程依据）。

## 它做什么

1. **抽取课程事实**：从 PPT / Word / PDF / Markdown / 教材 / 教案 / 培养方案 / 大纲中提取课程信息与教学内容，记录来源与证据角色。
2. **按通用契约编制教案**：每个教学单元写出教学目标（知识/能力/价值）、课程思政、专创融合、教学重难点、教学方法、教学用具、教学设计、教学过程（七步）、教学反思、参考文献。
3. **双输出**：工作 Markdown 便于协作与修改；Word 按该校教案模板结构编译（A4、封面「教学设计」、课程基本信息表、教学单元表）。

## 目录结构

```text
lesson-plan-compiler/
├── SKILL.md                                # 触发、路由与核心合同
├── agents/openai.yaml                      # Codex 接口（显示名、默认提示）
├── assets/
│   ├── lesson-plan-master.docx              # 样式模板：空白 Word 主文档（页面/字体/命名样式，渲染时加载）
│   ├── lesson-plan-working-template.md     # 数据合同：可复用工作 Markdown 模板
│   └── lesson-plan-working-example.md      # 虚构示例（不得当作真实事实）
├── references/
│   ├── data-template.md                    # 数据模板：字段项目名、填写特点、来源转化、注意事项
│   ├── language-style-rules.md             # 教案语言特点（跨课程蒸馏）
│   ├── word-layout-rules.md                # 样式模板的版式说明（页面/字体/表格结构）
│   └── markdown-review-rules.md            # 工作文档合同与审查级别
└── scripts/
    ├── make_master.py                      # 重新生成样式模板 lesson-plan-master.docx
    └── lesson_plan.py                      # check / render / inspect
```

## 用法

```powershell
# 0.（可选）重新生成样式模板（调整页面/字体/命名样式后）
py -3.12 scripts/make_master.py

# 1. 检查工作 Markdown（结构、必填字段、七步教学过程、占位符）
py -3.12 scripts/lesson_plan.py check --input 教案-working.md --report review.md --level basic

# 2. 编译 Word（加载 lesson-plan-master.docx 作为样式模板，先跑 basic 门禁）
py -3.12 scripts/lesson_plan.py render --input 教案-working.md --output 教案.docx

# 3. 终稿清洁门禁（批注 / 修订 / 隐藏文字 / 占位符 / AI 署名 / 工具元数据）
py -3.12 scripts/lesson_plan.py inspect --input 教案.docx --report inspection.md
```

要求 Python 3.10+、`python-docx`、`PyYAML`。Windows 下用 `py -3.12`（或 `python`）均可。

## 输出 Word 结构

渲染结果与「机械与动力工程学院」教案模板一致：

| 部分 | 结构 |
| --- | --- |
| 封面 | 「教 学 设 计」标题 + 7×2 信息表 + 「教务处制」+ 日期 |
| 课程基本信息 | 16×6 课程信息表 + 5 条填表说明 |
| 教学单元 | 每个单元一张 5 列 × 20 行表（含教学过程七步） |

## 模板三轴

一份标准教案由三条正交规则共同决定：

| 轴 | 文件 | 决定 |
| --- | --- | --- |
| 样式模板 | `assets/lesson-plan-master.docx`（渲染时加载）+ `references/word-layout-rules.md`（版式说明） | 排版与可见表单格式（页面、字体、表格边框） |
| 数据模板 | `references/data-template.md` | 每张表的字段项目名、填写特点、注意事项，以及 PPT / 教材 / 参考资料如何转化进各字段 |
| 语言规则 | `references/language-style-rules.md` | 措辞与句式 |

样式模板与 `course-syllabus-compiler` 的 `syllabus-master.docx` 同构：一份空白 Word 主文档承载页面设置与命名样式，渲染器加载它、清空正文后按数据重建；改样式只需改 master（或用 `make_master.py` 重新生成），不改渲染代码。

数据模板是「其他教学资料转化进来」的契约：它逐字段写明「来源角色 → 转化规则」，让大纲、培养方案、教材、PPT、参考资料各归其位，缺项标 `to_confirm` 而不猜测补齐。

## 防过拟合

语言与版式规则由多门课程教案**并集蒸馏**得到，区分 `stable-core`（多份一致）与 `conditional-variant`（按课程类型取舍）。规则只描述共同契约，不保存任何真实课程内容或教师身份。生成时逐条对齐契约，而不是复刻某一门课。

## 隐私

不提交课程包、原始教学材料、转换缓存、构建状态、生成产物、学生数据、本地路径或私人审查日志。公开资产与示例保持虚构并去身份化。
