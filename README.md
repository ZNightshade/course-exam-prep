# course-exam-prep · 课程考试资料整理

一个用于 [Claude Code](https://claude.com/claude-code) 的 Skill：把一门课的课件（PDF / PPT / Word 等）整理成期末能直接用的复习资料，并编译成可打印的 Word 手册。支持两种模式：**开卷**做"翻得快"的纸质速查手册，**闭卷**做"记得住 + 想得通 + 能应用"的复习材料。

课件多、上下文长，所以本 skill 靠**一个自包含的规则文件**统一口径，再用**多个独立 subagent**分讲整理（每讲一份独立 context），避免所有内容堆进同一个上下文里把对话撑爆。

---

## ✨ 它能做什么

- **支持多格式课件**：PDF / PPTX / DOCX / TXT / MD 都能处理。
- **公式忠实还原**：理工/数学课的公式在成品里正常显示，而不是字面 `$x$` 文本。
- **两种应试模式**：
  - **开卷** — 求全、宁多勿漏，做带索引、可快速翻查的纸质速查手册（总索引按拼音排序，像字典）。
  - **闭卷** — 覆盖**记忆 / 理解 / 应用**三层，含掌握分级、概念/推导卡、应用迁移题、自测题库；尤其算法/理工课不会做成纯背诵卡。
- **可打印成品**：自动生成排版好的 Word（封面、每讲另起一页、页眉显示讲名、表头跨页重复、页脚页码）。

---

## 🔄 工作流程

```
阶段0  建规则 + 试点第一讲（一个对话内一气呵成）
   问信息 → 读「试点讲」课件 → 分析题型/知识单元
   → 生成《规则》/汇总骨架/manifest → 紧接着把这一讲整理出来（试点）
   → 和用户过一遍、调到满意（踩的坑写进规则文件）

阶段1  派 subagent 整理剩余各讲（试点确认 OK 后才做）
   每讲一个独立 subagent（独立 context，防主线程爆炸）→ 各只产自己那讲的笔记
   → 串行合并：单一写入者把各讲并进 00_ 汇总、勾进度、记坑（零竞态）

阶段2  编译成品（一次）
   build_word.py 合并所有 md → 可打印的 Word 手册
```

---

## 📦 安装

```bash
pip install -r requirements.txt
```

依赖：`pymupdf` `python-pptx` `python-docx` `pypinyin` `latex2mathml`（`lxml` 通常随 latex2mathml 一并安装）。

### 环境要求

- **Windows 优先**：脚本在 Windows + 中文环境下测试。所有 python 调用前请加 `PYTHONIOENCODING=utf-8`，否则 Windows GBK 编码下中文会 `UnicodeEncodeError`。
- **公式渲染依赖 Office**：`build_word.py` 把 LaTeX 转 Word 原生公式时，需要本机 Microsoft Office 自带的 `MML2OMML.XSL`（脚本会自动在 `C:\Program Files*\Microsoft Office\...` 下查找）。找不到或未装 `latex2mathml` 时，公式会回退成字面文本——有公式的课请务必装好。
- **PPT 转 PDF 依赖 PowerPoint**：`pptx_to_pdf.py` 优先用本机 **PowerPoint COM**（保真最好），不可用时回退 **LibreOffice**。公式密集的 slide 课，建议在主线程一次性批量转好 PDF 再派 subagent（PowerPoint COM 是单实例，多个并行转会互相卡死）。

---

## 🚀 使用方法

在 Claude Code 对话里显式调用：

```
/course-exam-prep
```

然后按提示告诉它课程信息（课件路径、共几讲、开卷 / 闭卷）即可，Claude 会走完上面的三阶段流程。

---

## 📁 目录结构

```
course-exam-prep/
├── SKILL.md                          # skill 主文件：完整流程与原则
├── README.md
├── requirements.txt
├── references/
│   ├── extraction.md                 # 课件提取（文字 + 看图）详细手册
│   ├── open-book.md                  # 开卷：每讲模板、索引设计、编译、自检、坑
│   └── closed-book.md                # 闭卷：记忆/理解/应用三层模板、应用题、自测、坑
├── scripts/
│   ├── extract.py                    # 多格式提取 + 文字量统计 + 渲染图片
│   ├── pptx_to_pdf.py                # PPT→PDF（PowerPoint COM / LibreOffice）
│   └── build_word.py                 # 合并 md → 可打印 Word（拼音索引 / 自测遮挡 / 分页页眉页码）
└── assets/
    ├── rules_template.md             # 开卷规则模板
    ├── rules_template_closed.md      # 闭卷规则模板
    ├── manifest_example.json         # 开卷编译清单示例
    ├── manifest_example_closed.json  # 闭卷编译清单示例
    └── subagent_prompt_template.md   # 单讲 subagent 的派发指令模板
```

---

## 🔧 安装到 Claude Code

```bash
# 用户级（所有项目可用）
git clone https://github.com/ZNightshade/course-exam-prep.git ~/.claude/skills/course-exam-prep

# 项目级（仅当前项目）
git clone https://github.com/ZNightshade/course-exam-prep.git .claude/skills/course-exam-prep
```

---

## ⚖️ 开卷 vs 闭卷

| | 开卷 | 闭卷 |
|---|---|---|
| 目标 | 翻得快 | 记得住 + 想得通 + 能应用 |
| 产物 | 速查表 + 拼音排序总索引 | 掌握分级 + 概念/推导卡 + 应用题 + 自测题库 |
| 取舍 | 求全、宁多勿漏 | 覆盖记忆/理解/应用三层，不做纯背诵卡 |
| 编译 | 总索引按拼音排序、每讲分页 | 自测题另起页加"先做后看"提示、按讲顺序 |

建规则时会用提问确认本课程是开卷还是闭卷，再选对应模板——两者的取舍和产物完全不同，不会默认。
