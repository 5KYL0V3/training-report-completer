# 环境与运行

- 使用用户提供的实际文件路径；不内置用户名、桌面映射、盘符或历史会话路径。
- 脚本按当前工作目录解析相对路径，并支持 `~`；桌面重定向应由调用方提供真实路径。
- 仅在已知路径失败后作一次有针对性的定位，不默认递归搜索用户目录或整个磁盘。
- 工作副本、扫描缓存和 patch 保存在本次任务的工作目录，不读取其他任务的历史文件。

## 运行与渲染

Python 3.10 或更新版本；两个脚本需要 `lxml` 和 `python-docx`，不依赖 PyYAML。可用 `python -m pip install -r requirements.txt` 安装依赖；若当前环境已有依赖，直接复用。

视觉验收使用当前环境已提供且已验证的 DOCX 渲染工具。Windows 安装 Microsoft Word 时可用独立的隐藏 COM 实例只读导出 PDF；也可使用环境中已有的文档渲染工具。不要假设特定软件、安装路径或运行时存在，不结束用户已有的 Word 进程。无法渲染时如实说明视觉验收未完成。

## 失败上限

先区分页面未加载、定位错误、登录问题、权限限制和工具不可用，再修复对应原因。同一种工具或路径首次失败后，最多正常重试 1 次；再次失败后仅使用已知备用路径 / 方法一次，仍失败则停止该分支并报告。已确认同一失败原因时，不换表面调用方式继续尝试。确定的登录、验证码、权限或不存在错误无需盲重试；不得为确认同一事实继续遍历系统。沿用本机已验证的运行时和渲染路径；常规操作已获用户授权便继续，仍遵守实际工具权限机制。浏览器登录复用、动态图表就绪检查和早期阻塞验证见 report-rules 的“真实检索与截图”。

## 脚本调用与 patch

使用当前依赖工具返回的 Python，以 `-X utf8` 调用本技能脚本。以下路径相对于本技能目录，输入 / 输出文件使用实际绝对路径：

```text
python -X utf8 scripts/inspect_report.py report.docx manual.docx --output work/inspection.json
python -X utf8 scripts/apply_report.py report.docx work/patch.json
```

inspect 一次读取报告和可选 DOCX 指导书，stdout 输出紧凑 JSON，`--output` 缓存未截断全文。带 `truncated` 的段落需要原文时，仅从缓存读取对应 id，不重新扫描包。默认每段最多 280 字；只有按需调试才使用 `--debug` 输出 XML。PDF 指导书用现有 PDF 工具抽取一次。

位置 ID 为 `word/document.xml#tc3`（单元格）或 `#p12`（段落），从扫描结果直接取，不能猜编号。inspect schema 2 提供单元格的 `paragraph_ids`、子段落 / 嵌套单元格和表格的 `parent_cell`，截图占位仅按段落计数。嵌套表格填叶级单元格或指定段落，不整体替换父单元格。合并单元格按实际 XML 定位，不用界面列号硬算；blank、颜色只是候选提示。图片只做“存在”判断，真实来源由 agent 核实。

patch 示例（期望文字与 source_sha256 均来自本次扫描，不照抄示例值）：

```json
{
  "source_sha256": "本次扫描的摘要",
  "operations": [
    {"target": "word/document.xml#tc1", "field": "name"},
    {"target": "word/document.xml#tc3", "field": "student_id_last2"},
    {"target": "word/document.xml#tc5", "field": "date", "format": "YYYY.MM.DD"},
    {"target": "word/document.xml#tc7", "text": "一次集中生成的过程", "font_size_pt": 12},
    {"target": "word/document.xml#tc9", "expect": "（写）", "field": "reflection", "text": "第一段心得。\n第二段心得。", "font_size_pt": 12, "clean_prompt": true},
    {"target": "word/document.xml#tc11", "field": "teacher", "text": "从顶部提取的姓名"},
    {"target": "word/document.xml#tc13", "field": "teacher_comment", "text": "报告过程清晰，检索条件记录明确，能够结合结果说明文献筛选方法，书目信息整理基本完整。"},
    {"target": "word/document.xml#tc15", "field": "score"}
  ],
  "required": ["word/document.xml#tc1", "word/document.xml#tc7"]
}
```

- 字号由 agent 按 report-rules 判断：无特殊要求且空间适合的新增普通文字，在该条 `set_text` 操作传 `"font_size_pt": 12`。有明确字号要求可传对应磅值；需沿用模板则省略或传 null。脚本只设置新建文字 run 的 `w:sz` / `w:szCs`（12 pt = 24 半磅），保留字体名称及其他属性，不改表格或全文样式；不传参数仍继承模板。该参数仅支持 `set_text`（含教师评语和自动评分），`replace` 保留原混合格式，图注沿用现有插图逻辑，不对这些操作强制字号。
- 普通 `set_text` 的 `expect` 默认为空；非空区域须提供原文作前置条件，防止覆盖错误位置。`clean_prompt` 仅在确认该区域颜色 / 底纹为填写提示时使用。
- `field` 支持 name、student_id、class、major、college、school、student_id_last2、date、teacher、teacher_comment、reflection、score；身份默认来自正式 student-profile.md。缺必需字段则在写文件前报错，继续正文时可去掉依赖身份的操作，在 work 目录保存部分稿并说明缺项。
- `reflection` 仅用于心得/结论正文，`teacher_comment` 仅用于教师评语正文；两者 set_text 自动对每个自然段首行缩进 2 字符，text 中的换行在这两个字段内表示新段落，其他字段保持原换行行为。已有文字只改缩进可用 `{"op":"format_paragraphs","target":"扫描的正文段落或叶级单元格ID","field":"reflection"}`（或 teacher_comment），不传 text、font_size_pt 或 clean_prompt；它只改段落缩进，保留文字、图片和其他格式，跳过空段、编号列表和带图段落。不得把栏名、图注、姓名或签名作为 target，混合单元格应逐个选择正文段落。
- 课表使用当前任务资料或本机可选的 `references/timetable.local.md` / `references/timetable.local.json`，按 `references/timetable.md` 核对学期和课程后，将教师、日期等显式传入 patch。apply 不猜课程或周次基准；缺少资料时不追查无关个人信息。
- 教师评语、评分操作须在当前任务授权范围内；`teacher_comment` 传符合实际报告的正式文字；`score` 对空白栏可省略 text，脚本依据 required 缺项、待截图和 requirements 状态写“优”或“良”。patch 的可选 `requirements` 是 `{id,status,...}` 列表，status 仅取“原要求完成”“提供替代成果”“仍未完成”；后两者均使验收为 partial。需要实际导出、下载等的任务必须列出这些要求，脚本不能自行核实外部操作。所有新 text / caption 拒绝明确的审核或建议标签，不扫描清除已有真实批阅，也不禁用正常语义的“建议”。同义免责声明仍由 agent 在一次基础验收中核对。
- 日期优先在 patch 的 `date` 中传明确指定值，`date_format` 或每项 `format` 沿用模板。已有明确日期默认保留；没有明确日期时脚本取运行机器的本地当天，agent 确认这与用户时区一致；不一致时传用户本地日期。`local_date` 只用于这种时区补偿或固定日期回归测试，不把示例日期作为日常默认。
- 已有签名图片的教师区域自动跳过；普通有图区域不允许整体文字替换。`replace` 可对唯一的 `old` 子串做精确文字替换而保留图片；用于本轮重复姓名清理时，先确认写入记录，不删原稿已有文字。`keep` 保留并记录原因。
- 一个目标只传一个操作，不同时修改父单元格及其子段落。复杂内容控件、域、超链接、嵌套表格等整体替换会拒绝；嵌套表格使用扫描的叶级单元格 / 段落 ID，其他复杂位置改用精确 replace 或已有文档工具，不重新生成整个文档。
- `insert_image` 定点向已扫描段落插图，可加图注；现有图片和表格属性不重置，按 width_inches 等比缩放。必须先核实真实来源及对应要求，再传 evidence；verified 是 agent 的核验声明，脚本不抓取截图也不能验证网页真实性。有签名图片的指导教师区域（包括其子段落）自动保留。
- 有截图占位的普通文字替换会跳过并返回 `screenshot_slots_unverified` / partial；只有定点插入已核实图像才能移除对应 marker。未完成截图保留原占位，保留不等于完成。
- 默认输出到原目录，也可用 `--output` 指定 work 中间稿或正式交付目录。替换文件名身份占位符、清除末尾下载副本标记，并去掉 `+`；显式输出名也去加号，使用返回的 output 作为实际链接。无变化则加 `_已填写` 防止覆盖原件。若数字括号是课次 / 版本，或模板有特殊命名要求，传 `--output` 明确指定；去加号后仍校验目标碰撞，绝不覆盖原件或其硬链接。`--overwrite-generated` 仅用于已确认的以前生成版本，不可用于来源不明文件。
- 返回 `opens`、缺项、跳过原因和待截图项即为一次结构基础 QA。`visual_qa=not_performed` 需补一次已知方法的视觉检查；这不是再跑多套结构校验的理由。工具报错先看具体原因，按上方失败上限执行。

定点插图操作示例（放入 operations；ID、路径、来源和图注均换为本次已核实材料）：

```json
{
  "op": "insert_image",
  "target": "word/document.xml#p12",
  "expect": "（截图）",
  "marker": "（截图）",
  "image": "screenshots/result.png",
  "width_inches": 4.8,
  "caption": "图1 指定数据库检索结果",
  "evidence": {
    "verified": true,
    "source": "本次实际访问并核实的页面 URL",
    "requirement": "截图1：指定数据库检索结果"
  }
}
```

image 相对路径以 patch 所在目录为基准。插入到目标段落末尾，图注作为紧随其后的段落；保留原有段落样式，不修改表格宽度、行高、跨页属性。宽度应取模板可用宽度，出现实际问题再局部修复。仅在真实检索或软件取证要求下使用这些截图能力。

维护本技能时用 `python -X utf8 -m unittest discover -s tests -v` 运行合成文档回归测试；不使用真实用户原件，也不跑完整浏览器实训。
