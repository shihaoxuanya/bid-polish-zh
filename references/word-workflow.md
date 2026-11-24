# Word 文件操作与验收

在用户选定的原稿旁预留一个新的修订文件名。索引、修改计划、回执、接受稿和页面图都放在本次任务独立的临时目录。用户明确要求提供净稿时再交付接受稿，并说明仍待复核的批注和既有修订。示例中的 `$py` 为带有 lxml 的 Codex 文档运行时，`$scripts` 为本技能的 scripts 目录，`$qa` 为已创建的临时目录。使用文档工具时先解析本机工作区依赖。PowerShell 示例以支持 UTF-8 的 pwsh 为准。

## 读取和拟改

```powershell
& $py -X utf8 "$scripts/bid_docx.py" inspect $original --out "$qa/index.json"
```

按 index.json 中的段落 ID、章节、样式、表格标志、保护原因和可见文字逐章审阅。保护提示描述 Word 结构；还需从内容辨认采购方引文、偏离表等固定表述。标题、页眉页脚、文本框、域、签章以及表格结构均无需重做。把跨多种格式的长句缩到一个确实需要编辑的短语。对已经有修改或批注的文字，优先在相邻安全位置留下建议。

用 UTF-8 写 plan.json。将示例指纹换为本次索引中的完整 SHA-256，ID 和原句都从该索引取。每条 before 或 quote 在目标段内应只出现一次，相邻修改不可重叠。单独删除一个词组时 after 允许为空字符串。数组可以留空。

对本轮拟改段落，把必须逐字保留的项目用语列入 `protected_terms`。每项写段落 ID、原稿确实出现的连续原词，以及来源文件和章节或评分项。来源只在本稿时写本稿位置。相同术语关联多个响应段时逐段填写。原稿未采用条款原词且是否等义存疑时，使用 `comments` 说明两处措辞。没有需要逐字保护的项目用语时，可省略 `protected_terms`。

```json
{
  "source_sha256": "<本次原稿指纹>",
  "protected_terms": [
    {"p":"P0032", "term":"电话响应", "source":"评分表第4项“电话响应”"}
  ],
  "edits": [
    {"p":"P0014", "before":"对恢复演练工作进行组织实施", "after":"组织恢复演练"}
  ],
  "comments": [
    {"p":"P0032", "quote":"30分钟内电话响应", "text":"技术响应第3.2节写30分钟，采购需求第4项写15分钟。请确认电话响应时限。"}
  ]
}
```

每条提案应能独立接受或拒绝。若同段有多处改动，先分别读只接受一处的结果，再读全部接受的结果。比如把时限从前句挪到后句，会产生两条看似只删字、加字的修订，却改变了受时限约束的动作。遇到这种情况保留原承诺；需要提示时在原句旁写清哪一项有疑问。不要用全章替换或把 XML 导出成纯文本再重建标书。

## 生成和对照

```powershell
$receipt = & $py -X utf8 "$scripts/bid_docx.py" apply $original "$qa/plan.json" --out $tracked
if ($LASTEXITCODE -ne 0) { throw '修订未生成，请检查原稿和计划。' }
[IO.File]::WriteAllText("$qa/receipt.json", ($receipt -join [Environment]::NewLine), [Text.UTF8Encoding]::new($false))
& $py -X utf8 "$scripts/verify_docx.py" $original "$qa/plan.json" $tracked "$qa/receipt.json" --accepted "$qa/expected.docx"
if ($LASTEXITCODE -ne 0) { throw '结构或正文不一致，请核对后重新生成。' }
```

apply 按原稿中的位置核对提案，以本轮唯一的审阅人署名写入 Word 原生修订和批注。拟改处会改掉已登记的项目原词时，脚本保留该处原文，在回执和批注中写出原词及登记的依据，同段其他合适的轻改仍可进行。脚本另行比对可识别的指标、书面批准条件、义务或优先安排、故障时限的明确起算事件及相应动作。术语是否来自对应评分项、跨段代词指谁，以及隐含的起算点，均须对照项目材料判断。

读取回执的每个状态。涉及域、锚点或混合格式的提案可能转换为批注；若回执提示原句不存在或出现多次，重新读取原稿并定位，不能将落在段末的建议算作成功改稿。校验失败时保留输入，修正计划或脚本，用新文件名重试。不可用已有原稿、修订版或已存在的验证文件作为输出目标。

verify_docx 按每条提案的段落和原稿字符位置核对修订编号，检查单独接受该处修改是否得到拟改文字，再校验全文接受或拒绝后的结果。普通原位批注须实际圈中指定引文；因域或混合格式回退的批注则核对实际落点和说明。校验同时覆盖字符格式、图像、页眉页脚、段落属性及表格结构。expected.docx 是独立核对用的临时接受稿，保留既有审阅记录。

## 在 Word 中复核

Windows 上安装 Microsoft Word 时，使用 Windows PowerShell 运行自带的复核脚本，OutDir 指向尚不存在的目录：

```powershell
$batch = Get-Content -Encoding utf8 -Raw "$qa/receipt.json" | ConvertFrom-Json
powershell.exe -NoProfile -File "$scripts/word_review.ps1" -Original $original -Tracked $tracked -Author $batch.author -ExpectedAccepted "$qa/expected.docx" -OutDir "$qa/word"
if ($LASTEXITCODE -ne 0) { throw 'Word 复核未通过，请检查 review.json。' }
```

脚本在测试副本中按作者接受或拒绝本轮修订，分别核对原稿与预期正文；在接受稿刷新目录和分页，生成审阅结果和正文 PDF。查看 review.json 及原稿与接受稿的版面，实际交付的修订文件仍由用户决定接受或拒绝。若需交付只接受本轮修改的版本，检验临时接受稿的字段与批注后再复制到新文件。发现个别目录域缓存过期时，告诉用户接受修订后在 Word 中更新目录。

使用工作区捆绑的文档渲染器，或将以上由 Word 导出的 PDF 用 Poppler 转为逐页 PNG。按隐藏修订标记的视图检查全部页面及修改位置，留意表格跨页、末行与页脚、题注和图片位置。正文自然换行换页可以保留。页脚显示的页码可能与 PDF 物理页序不同，按索引实际采用的编号核对。独立索引若无超链接，按“条款号＋响应标题＋该条正文”定位到原页、现页后给出对照；只有相同标题而无法确认时注明未定位。不能仅由总页数推断所有索引正确。

若当前环境没有可用的 Word，仍可完成结构和文字校验并在可用的受支持渲染环境检查版面。交付说明要区分已完成的验证与尚未实测的 Word 行为或分页。加密、受保护或已签名且无法安全写入的原稿，保留文件，给出能够定位的文字建议。
