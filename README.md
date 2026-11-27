# Nature EndNote Citation

把 Markdown 或 Word 论文稿中的临时编号引文转换为 **EndNote Cite While You Write（CWYW）字段**。输出是 `.docx`，可在安装了 EndNote CWYW 的 Microsoft Word 中更新引文、生成参考文献表，并应用所选的输出样式。

[English quick start](#english-quick-start)

## 适用场景

- **Markdown 转 Word**。将 `[1]`、`[4,6]`、`[23-28]` 等编号映射到对应的参考文献，并在生成 Word 时嵌入引文字段。
- **已有 Word 稿件**。把正文里的临时引文替换为 EndNote 字段，同时保留标题、题注和稿件结构。
- **投稿前检查**。用脚本检查 `.docx` 中是否存在引文字段、参考文献表字段和嵌入的文献记录，再在 Word 中验证排版。

技能以 `ADDIN EN.CITE` 保存正文引文及其文献元数据，以 `ADDIN EN.REFLIST` 标记参考文献表的位置。组合编号或连续编号可以在一个引文字段中容纳多条文献记录。

## 安装到 Codex

需要支持本地技能的 Codex 和 Git。以下命令把仓库克隆到用户级技能目录。在 Windows 上使用 PowerShell，在 macOS 或 Linux 上使用终端。

**Windows PowerShell**

```powershell
git clone https://github.com/zhuozhaowu/nature-endnote-citation.git "$HOME/.agents/skills/nature-endnote-citation"
```

**macOS / Linux**

```bash
git clone https://github.com/zhuozhaowu/nature-endnote-citation.git "$HOME/.agents/skills/nature-endnote-citation"
```

如果只想在一个项目中使用，从项目根目录把克隆目标改为 `.agents/skills/nature-endnote-citation`。安装目录下应直接包含 `SKILL.md`。已有同名技能时，可以复用原安装，避免技能列表出现重复项。如果 Codex 没有显示该技能，重启后再查找。 [Codex 技能目录与发现规则](https://learn.chatgpt.com/docs/build-skills#where-codex-loads-local-skills)

## 使用方法

1. **准备稿件和参考文献。** 提供 `manuscript.md` 或 `manuscript.docx`，以及从 EndNote 导出的 `references.enw`。如果导出顺序未与稿件编号对齐，再提供编号与文献的对应表，例如用 DOI 或完整标题标识每条文献。有期刊模板或含有正常 CWYW 字段的样例 `.docx` 时，可以一起提供。
2. **在 Codex 中调用技能。** 把稿件和参考文献放到任务可访问的位置，然后复制下面的提示词。需要指定显示样式时，直接写明，例如「保留方括号，上标显示」。
3. **在 Word 中更新。** 用启用了 CWYW 的 Microsoft Word 桌面版打开生成的 `.docx`。在 **EndNote** 选项卡选择相应的 **Style**，点击 **Update Citations and Bibliography**。检查正文引文与文献表能否刷新，再核对引用顺序、组合引文、缺项和标点。 [EndNote 官方排版说明](https://docs.endnote.com/docs/endnote/2025/windows/v1/content/09word/basic_instructions1.htm)

```text
使用 $nature-endnote-citation 将 manuscript.md 和 references.enw 转成
manuscript-endnote.docx。按提供的编号对应表匹配文献，把 [1]、[4,6]、
[23-28] 转成 EndNote CWYW 字段，并在 References 标题下插入文献表字段。
保留标题、题注和原有编号显示。遇到无法唯一匹配的编号时，请列出缺项，
不要猜测。交付前运行字段验证脚本并报告结果。
```

仓库提供 Codex 工作流、基于 `python-docx` 的字段生成示例，以及可单独运行的检查脚本。文档转换由 Codex 根据稿件和参考文献执行，最后的参考文献样式由 Word 中的 EndNote 输出样式决定。

## 用命令行检查生成的 DOCX

在仓库根目录执行。验证脚本只用 Python 标准库，可在未安装 Word 或 EndNote 的环境中运行。

```bash
python scripts/validate_endnote_fields.py "manuscript-endnote.docx" --min-cites 1 --require-reflist
```

| 参数或输出 | 含义 |
| --- | --- |
| `--min-cites 1` | 要求正文至少检测到 1 个 `ADDIN EN.CITE` 标记，默认值为 1 |
| `--require-reflist` | 要求检测到至少 1 个 `ADDIN EN.REFLIST` 标记 |
| `embedded_records` | `word/document.xml` 中检测到的 `<record>` 标记数 |
| `visible_numeric_tokens_in_xml` | XML 中匹配到的编号文本数，可能包含字段内容或其他编号 |
| 退出状态 `0` / `1` | 检查通过 / 未通过 |

脚本只读取文档正文的 XML 并做标记计数。引文与文献是否匹配、字段能否由 EndNote 更新，以及最终排版是否正确，需在 Word 中核对。

## 常见问题

| 现象 | 排查方法 |
| --- | --- |
| Word 没有 **EndNote** 选项卡 | 检查是否使用兼容的 Word 桌面版，以及 CWYW 加载项是否已安装并启用。 |
| 引文可见，但脚本报 `ADDIN EN.CITE` 为 0 | 检查稿件是否只保留了编号文字，或是否被其他编辑器另存后丢失了字段。 |
| 脚本通过，但文献表空白或样式不符 | 在 Word 中选择正确的 EndNote 样式并运行更新。如果 EndNote 提示找不到记录，打开对应的 EndNote 库，并核对编号与文献的映射。 |
| Word 更新后出现两份文献表 | 检查原稿是否保留了手工录入的参考文献列表。先核对引文字段生成的版本，再清理重复内容。 |

Word 中的 CWYW 字段可携带引用记录（Traveling Library）。协调作者修改时，保留含字段的 `.docx` 主稿，并用支持 CWYW 的 Word 编辑。 [EndNote 字段说明](https://docs.endnote.com/docs/endnote/2025/windows/v1/content/09word/field_codes.htm) · [Traveling Library](https://docs.endnote.com/docs/endnote/2025/windows/v1/content/09word/the_traveling_library.htm)

## 仓库文件

- [SKILL.md](SKILL.md) 包含转换流程、Word 字段结构和 Python 示例。
- [agents/openai.yaml](agents/openai.yaml) 定义技能在 Codex 中的显示信息和默认提示词。
- [scripts/validate_endnote_fields.py](scripts/validate_endnote_fields.py) 检查 DOCX 正文中的 EndNote 字段标记。

## English quick start

This Codex skill converts citation placeholders in Markdown or Word manuscripts into **EndNote CWYW fields**. It embeds cited reference metadata and adds an EndNote bibliography field.

1. Clone the repository into `$HOME/.agents/skills/nature-endnote-citation` using the [installation command](#安装到-codex).
2. Give Codex your manuscript, an EndNote `.enw` export, and a citation-number mapping if the export order differs from the manuscript.
3. Ask Codex to create an editable Word document with `$nature-endnote-citation`, then run the validator from the repository root.

```text
Use $nature-endnote-citation to turn manuscript.md and references.enw into
manuscript-endnote.docx. Preserve headings, captions, and citation labels.
Use the provided number-to-reference mapping, report unmatched citations, and
run the validator on the resulting DOCX.
```

```bash
python scripts/validate_endnote_fields.py "manuscript-endnote.docx" --min-cites 1 --require-reflist
```

Open the output in desktop Word with the EndNote CWYW add-in, choose an output style, and select **Update Citations and Bibliography**. The validator checks field markers in the document XML. Verify reference matching and formatting in Word. [EndNote documentation](https://docs.endnote.com/docs/endnote/2025/windows/v1/content/09word/basic_instructions1.htm)
