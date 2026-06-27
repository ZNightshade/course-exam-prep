# -*- coding: utf-8 -*-
"""把 PPT/PPTX 整套**忠实**导出为 PDF，用于渲染公式页/版式图。

为什么要它：`extract.py --images` 对 pptx 只能导出**内嵌图片 blob**，
**抓不到公式对象（OMML 数学公式）和整页版式**，而且常因坏图崩溃。很多课件的
公式/矩阵/几何示意是公式对象或排版出来的，纯靠 extract.py 会整页漏掉。
正确做法是先把 pptx 转成 PDF（公式按渲染后的样子落进 PDF），再用
`extract.py "x.pdf" --images ...` 整页渲染成 PNG 看图。

转换优先用本机 **PowerPoint COM**（保真最好），不可用时回退 **LibreOffice**。

用法：
    python pptx_to_pdf.py "课件/第02讲.pptx"                 # 生成 课件/第02讲.pdf
    python pptx_to_pdf.py "课件/第02讲.pptx" --out tmp/x.pdf  # 指定输出
然后：
    PYTHONIOENCODING=utf-8 python extract.py "课件/第02讲.pdf" --images "img/第02讲" --lowtext-only
    # 用 Read 工具逐张看 PNG，把公式/示意图补进笔记；看完可删临时 PDF
"""
import io
import os
import subprocess
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')


def via_powerpoint(src, dst):
    """Windows + 已装 PowerPoint：用 COM 导出 PDF（保真最好）。"""
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    # 直接驱动 PowerPoint COM；SaveAs 格式 32 = ppSaveAsPDF
    ps = (
        "$p=New-Object -ComObject PowerPoint.Application;"
        f"$d=$p.Presentations.Open('{src}',$true,$false,$false);"
        f"$d.SaveAs('{dst}',32);$d.Close();$p.Quit()"
    )
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, text=True)
    if r.returncode == 0 and os.path.exists(dst):
        return True
    sys.stderr.write((r.stderr or "PowerPoint COM 失败") + "\n")
    return False


def via_libreoffice(src, dst):
    """无 PowerPoint 时回退 LibreOffice（soffice --headless）。公式保真可能略差。"""
    src = os.path.abspath(src)
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    for exe in ("soffice", "libreoffice"):
        try:
            r = subprocess.run([exe, "--headless", "--convert-to", "pdf",
                                "--outdir", outdir, src],
                               capture_output=True, text=True)
        except FileNotFoundError:
            continue
        # soffice 按源文件名生成 <stem>.pdf，必要时改名到 dst
        gen = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
        if r.returncode == 0 and os.path.exists(gen):
            if os.path.abspath(gen) != os.path.abspath(dst):
                os.replace(gen, dst)
            return True
    return False


def main():
    argv = sys.argv[1:]
    args = [a for a in argv if not a.startswith('--')]
    if not args:
        print(__doc__); return
    src = args[0]
    dst = argv[argv.index('--out') + 1] if '--out' in argv else \
        os.path.splitext(src)[0] + ".pdf"
    if not os.path.exists(src):
        print(f"[错误] 找不到文件：{src}"); sys.exit(1)

    ok = via_powerpoint(src, dst)
    if not ok:
        print("PowerPoint COM 不可用，尝试 LibreOffice…")
        ok = via_libreoffice(src, dst)
    if not ok:
        print("[失败] 既无 PowerPoint COM 也无 LibreOffice。请手动在 PowerPoint 里"
              "「另存为 PDF」，再用 extract.py 按 PDF 流程渲染。")
        sys.exit(2)
    print(f"SAVED: {dst}")
    print(f"下一步：PYTHONIOENCODING=utf-8 python extract.py \"{dst}\" "
          f"--images \"img/<第N讲>\" --lowtext-only，再用 Read 逐张看图。")


if __name__ == '__main__':
    main()
