# 课件提取手册：两条腿走路（文字 + 看图）

整理任何一讲，第一步都是把课件内容**完整**拿到手。课件常是 PDF/PPT，里面
混着机器文字和图片。纯文字提取又快又省，但**图片里的内容抽不出来**——乐谱、
示意图、表格截图、扫描页、画作、地图等。漏了这些就会出错。所以固定两步：

## 第一步：抽文字 + 看文字量统计

```bash
cd "<课程目录>"
PYTHONIOENCODING=utf-8 python <skill>/scripts/extract.py "课件/第N讲.pdf"
```

- `PYTHONIOENCODING=utf-8` 必加。Windows 默认 GBK，中文会 `UnicodeEncodeError`。
- 输出末尾有「文字量统计」，逐页列出文字字数与图片数，并把**文字量 < 40 字**
  的页标为「⚠️建议看图」——这些就是内容主要在图里的页。

只想先看哪些页是图片页：
```bash
... python <skill>/scripts/extract.py "课件/第N讲.pdf" --scan
```

## 第二步：把图片页渲染成 PNG，再"看图"补全

对被标记的低文字页（或你判断含重要图示的页），渲染成图片，然后**用 Read 工具
逐张查看**，把图里的关键信息（人名、作品名、年代、乐谱特征、流程图要点）补进笔记。

```bash
# 只渲染文字少的页（聚焦、省资源）
... python <skill>/scripts/extract.py "课件/第N讲.pdf" --images "img/第N讲" --lowtext-only
# 或渲染全部页（课件整体很"图片化"时）
... python <skill>/scripts/extract.py "课件/第N讲.pdf" --images "img/第N讲"
```

然后对每个 PNG 调用 Read 工具查看（模型能直接读图）。这一步保证乐谱、图表、
扫描内容不被漏掉。

## 各格式注意事项

- **PDF**：pymupdf 直接抽文字、数图、渲染页。最省事。
- **PPTX**：抽文本框、表格、**演讲者备注**（备注里常有补充讲解，别漏）。
  `--images` 只能导出内嵌图片 blob，**抓不到公式对象（OMML 数学公式）和整页
  版式**，且老式/坏图会让旧版脚本中途崩（已改为跳过坏图）。所以——
  **凡是有公式/矩阵/几何示意图的 PPT，别指望 `extract.py --images` 拿到公式**，
  改走下面的"PPT→PDF→PNG"可靠流程：

  ```bash
  # 1) 整套 pptx 忠实转 PDF（公式按渲染样子落进 PDF）；优先 PowerPoint COM，
  #    无则回退 LibreOffice
  python <skill>/scripts/pptx_to_pdf.py "课件/第N讲.pptx"   # 生成 课件/第N讲.pdf
  # 2) 对 PDF 整页渲染成 PNG（PDF 路径能整页渲染，pptx 路径不能）
  PYTHONIOENCODING=utf-8 python <skill>/scripts/extract.py "课件/第N讲.pdf" \
      --images "img/第N讲" --lowtext-only
  # 3) 用 Read 工具逐张看 PNG，把公式/示意图补进笔记；看完可删临时 PDF
  ```
  公式抄进笔记/表格时注意：**表格单元格里的 LaTeX 不能含字面 `|`**，绝对值用
  `\lvert x \rvert`、范数用 `\lVert x \rVert`（含 `|` 会把表格列拆裂）。

  ⚠️ **公式/图密集的 slide 课，PPT→PDF 要在主线程"批量、提前"做完，别让 subagent 各自转**：
  `pptx_to_pdf.py` 走的 PowerPoint COM 是**单实例自动化服务**，多个并行 subagent
  同时调它转 PDF 会**互相卡死**（Word/PowerPoint COM 挂起、进程僵住）。所以阶段1 派
  subagent 之前，**主线程先在一次 PowerPoint 会话里把本课所有 PPT 串行转成同名 `.pdf`
  放课程目录**，再派 subagent——subagent 只用现成 PDF 渲染看图（pymupdf，纯 Python、
  并行安全），**指令里明确禁止它再调 pptx_to_pdf.py / PowerPoint COM**。
  主线程批量转的最省事写法是一次 COM 会话循环所有文件：
  ```powershell
  $pp = New-Object -ComObject PowerPoint.Application
  foreach ($f in $files) {
    $d=$pp.Presentations.Open($src,$true,$false,$false); $d.SaveAs($dst,32); $d.Close()
  }
  $pp.Quit()
  ```
  （注意：这是"按内容需要才转 PDF"——文字为主的课直接抽文字即可，**不要无脑把所有
  类型都转 PDF 看图**，那样慢、费 token 又易看错；转 PDF 只为拿公式/图。）
- **旧版 .ppt/.doc**：`extract.py` 抽文字会崩（python-pptx 不支持 .ppt 二进制格式）。
  但 PowerPoint COM 能正常把 .ppt 转 PDF——所以遇 .ppt **别纠结文字提取，直接走上面
  的 PPT→PDF→PNG 看图流程**；.doc 则提示用户另存为 .docx。
- **音频/视频**：一律不抽取、不试听。听辨题的曲名/答案课件文字里通常已写明，
  照文字整理即可。课件里若有大视频（如内嵌 mp4），忽略它，只处理文字与图。

## 依赖

```bash
pip install pymupdf python-pptx python-docx pypinyin
```
（pymupdf=PDF，python-pptx=PPT，python-docx=Word，pypinyin=索引排序）
