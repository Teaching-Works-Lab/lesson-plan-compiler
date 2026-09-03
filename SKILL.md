---
name: lesson-plan-compiler
description: Use when a university course teaching plan (教案) needs to be planned, drafted, continued, reviewed, revised, checked, or compiled from course materials, a course syllabus (大纲), and optional training-program evidence into an editable Markdown working document and a standards-compliant Word document.
---

# 教案编制器（Lesson Plan Compiler）

从课程材料、课程大纲（和可选的培养方案依据）编制、续写、审查、校验一门课程的教案，产出**可编辑 Markdown** 与**规范 Word** 两份产物。路由到用户请求的那一份产物，不强制跑固定流水线。

## 条件路由

- 用户只要**通用 Markdown 模板**：返回可复用模板即停，不创建课程包、不生成 Word。
- 用户给了**原始课程材料 / 大纲 / 培养方案依据**：先抽取与课程事实，边写边在 `# 编制信息` 记录来源、证据角色、缺项与冲突；只生成工作 Markdown，不自动编译 Word。
- 用户要**续写、审查、修订**已有 `教案-working.md`：按 `basic / changed / enhanced` 级别检查受影响范围，不触发无关的全面重审。
- 用户要**编译 Word**：对有效课程包先跑 `basic` 门禁，再编译并执行清洁门禁；复用未变化的构建状态。
- 只审查或只检查**指定产物**：警告不升级为无关的全面审计。

## 模板三轴

一份标准教案由三条正交的规则共同决定，缺一不可：

1. **样式模板**：可复用的空白 Word 主文档 `assets/lesson-plan-master.docx`，决定排版与可见表单格式（页面、字体、表格边框）；渲染器加载它、清空正文后按数据重建。其版式规则记录在 [word-layout-rules.md](references/word-layout-rules.md)，可用 `scripts/make_master.py` 重新生成；
2. **数据模板**（[data-template.md](references/data-template.md)）决定每张表的项目名称、填写特点、注意事项，以及 PPT / 教材 / 参考资料等如何转化进各字段；
3. **语言规则**（[language-style-rules.md](references/language-style-rules.md)）决定措辞与句式。

## 核心合同

- 保留原始文件；区分 `official / provided / derived / reference_only / to_confirm / not_applicable` 六种来源状态，不得把参考材料升格为官方课程事实。
- 长文本（教学目标、课程思政、专创融合、教学重难点、教学反思）写成可读的 `####` 小节；短事实（章节、周次、课时安排、教学方法、教学用具、教学设计、参考文献）放在字段—值表；教学过程用固定三列表格，单元格内用 `<br>` 换行。
- 不逐字套用某一份样例教案：模板是全体教案的共同契约，不是单个案例的复刻。按 [language-style-rules.md](references/language-style-rules.md) 的稳定模式生成，条件性写法按课程实际类型取舍。
- 正式 Word 不得包含 AI/工具署名、批注、修订、隐藏文字、未解决占位符、来源状态标签或内部审查意见。
- 不要求固定流程；未请求 Word 时不生成 Word。仅模板、仅分析、仅审查的请求在对应产物通过门禁后即停。

## 资源

- 需要通用或课程专用工作 Markdown 时，复制 [lesson-plan-working-template.md](assets/lesson-plan-working-template.md)；用 [lesson-plan-working-example.md](assets/lesson-plan-working-example.md) 只理解合同，其数据为虚构，不得当作真实课程事实。
- 确定每张表填哪些字段、怎么填、从哪来、注意什么时读 [data-template.md](references/data-template.md)；生成或修订教学内容时读 [language-style-rules.md](references/language-style-rules.md)；编译或视觉复核 Word 前读 [word-layout-rules.md](references/word-layout-rules.md)；建立、修订或检查 Markdown 时读 [markdown-review-rules.md](references/markdown-review-rules.md)。
- 运行 `scripts/lesson_plan.py check --help` 做确定性结构检查；复用脚本，不写课程专用检查器。
- 编译 Word 用 `scripts/lesson_plan.py render --help`（默认加载 `assets/lesson-plan-master.docx` 作为样式模板，可用 `--template` 覆盖）；终稿用 `scripts/lesson_plan.py inspect --help` 检查清洁门禁。
- 需要调整页面、字体或命名样式时，编辑并重新生成样式模板：`py -3.12 scripts/make_master.py`。

## 编译命令

```powershell
py -3.12 scripts/lesson_plan.py check --input 教案-working.md --report review.md --level basic
py -3.12 scripts/lesson_plan.py render --input 教案-working.md --output 教案.docx
py -3.12 scripts/lesson_plan.py inspect --input 教案.docx --report inspection.md
```
