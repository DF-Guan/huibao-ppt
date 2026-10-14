# -*- coding: utf-8 -*-
"""export_preview — 一键导出 PPTX 为 PNG(验收第四阶段的一公里)。

优先 PowerPoint COM(高保真);无 PowerPoint 时尝试 LibreOffice;都没有则给出提示。

用法:
    python export_preview.py <成品.pptx> [导出目录]
"""
import os
import shutil
import subprocess
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

def export_com(pptx, out_dir):
    """PowerPoint COM 导出(Windows + Office 已装时最高保真)。"""
    try:
        import win32com.client  # noqa
    except ImportError:
        return False
    try:
        import win32com.client
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        deck = ppt.Presentations.Open(pptx, True, False, False)
        os.makedirs(out_dir, exist_ok=True)
        deck.SaveCopyAs(os.path.join(out_dir, "slide"), 18)  # 18 = ppPNG
        deck.Close()
        ppt.Quit()
        return True
    except Exception as e:
        print(f"COM 导出失败: {e}")
        return False


def export_lo(pptx, out_dir):
    """LibreOffice 兜底。"""
    for exe in ("soffice", "libreoffice"):
        path = shutil.which(exe)
        if path:
            os.makedirs(out_dir, exist_ok=True)
            r = subprocess.run([path, "--headless", "--convert-to", "png", "--outdir", out_dir, pptx],
                               capture_output=True, timeout=300)
            return r.returncode == 0
    return False


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    pptx = os.path.abspath(sys.argv[1])
    out_dir = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(pptx), "preview")
    if not os.path.exists(pptx):
        print(f"文件不存在: {pptx}")
        sys.exit(1)
    if export_com(pptx, out_dir) or export_lo(pptx, out_dir):
        # PowerPoint 会在目标目录下建以文件名命名的子目录,递归计数
        n = sum(len([f for f in fs if f.lower().endswith((".png", ".jpg"))])
                for _, _, fs in os.walk(out_dir))
        if n == 0:
            print(f"导出命令执行但未找到 PNG,请检查: {out_dir}")
            sys.exit(3)
        print(f"导出成功: {out_dir} ({n} 张)")
        sys.exit(0)
    print("未找到 PowerPoint 或 LibreOffice,请用 Office/WPS 手动导出 PNG 验收。")
    sys.exit(2)


if __name__ == "__main__":
    main()
