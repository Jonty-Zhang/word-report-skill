# word-report

写课程类 Word 交付物（结课论文、小论文、实验报告、课程汇报稿、读书报告、小组项目报告等）时用的 skill。目标是让 agent 写出来的文档像一个认真的学生自己写的，交上去不用再改一遍。

## 它管什么

- **字号**：全文任何位置不小于 10 pt，表格、图题、代码、页码都算。
- **图片**：按类型给宽度区间，上限 15 cm 宽、18 cm 高，保持比例，清晰度不低于 150 dpi。
- **缺点表述**：只在一处集中写，2–4 句，具体、向前看；其他章节不加让步尾巴。
- **禁用表述**："按照用户要求""课堂要求""以下是""更好的做法是…但为了简便…"这类元话语和自贬一律不出现。
- **自称与语气**：论文用"本文/笔者"，汇报稿用"我/我们"，实验报告少自称，全文一套。
- **排版**：无模板时按 `references/format-spec.md`（A4、2.5 cm、宋体五号、1.5 倍行距、标题层级……），有模板严格按模板。
- **去 AI 味**：`references/wording.md` 里有面向学术文体的 humanizer 清单。

## 文件

```
SKILL.md                     主流程、硬规则速查、交付前检查
references/format-spec.md    排版规范全文（含图片尺寸表）
references/wording.md        禁用表述、缺点写法、去 AI 味
references/report-types.md   各文体的自称、语气、结构
scripts/check_docx.py        成品检查：字号、图片、禁用短语、缺点分布、自称、空段、页面
scripts/fit_images.py        修正 docx 里超尺寸 / 变形的图片
scripts/docx_helpers.py      python-docx 建文档的小工具（字体、插图配图题、表格、代码块）
scripts/selftest.py          自测
```

## 安装

```bash
git clone https://github.com/Jonty-Zhang/word-report-skill.git
ln -s "$(pwd)/word-report-skill" ~/.agents/skills/word-report
ln -s ../../.agents/skills/word-report ~/.claude/skills/word-report
```

依赖：Python 3 + `python-docx` + `Pillow`。

## 用法

对 agent：写报告时会自动触发；也可以直接说"用 word-report 写……"。

手动检查一份文档：

```bash
python3 scripts/check_docx.py 报告.docx
python3 scripts/fit_images.py 报告.docx --inplace
```

自测：

```bash
python3 scripts/selftest.py
```
