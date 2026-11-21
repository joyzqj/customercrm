#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
客户通 CustomerCRM — 单文件构建脚本
把 index.html + css/app.css + 全部 js 内联为一个可离线双击打开的 HTML，
输出到 dist/客户通-外贸客户管理系统.html。

用法：python build_single_file.py
无第三方依赖，Python 3.8+。
"""
from __future__ import annotations

import datetime
import hashlib
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent
OUT_NAME = "客户通-外贸客户管理系统.html"


def read_text(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


def guard_script(js: str) -> str:
    # 避免内联时提前闭合 <script>
    return js.replace("</script>", "<\\/script>")


def main() -> int:
    index_path = ROOT / "index.html"
    if not index_path.exists():
        print("找不到 index.html", file=sys.stderr)
        return 1

    html = index_path.read_text(encoding="utf-8")
    inlined: list[str] = []

    def css_repl(match: re.Match[str]) -> str:
        href = match.group(1)
        css_path = ROOT / href
        if not css_path.exists():
            return match.group(0)
        inlined.append(href)
        css = css_path.read_text(encoding="utf-8")
        return f"<style>\n/* {href} */\n{css}\n</style>"

    def js_repl(match: re.Match[str]) -> str:
        src = match.group(1)
        js_path = ROOT / src
        if not js_path.exists():
            return match.group(0)
        inlined.append(src)
        js = read_text(src)
        return f"<script>\n/* {src} */\n{guard_script(js)}\n</script>"

    html = re.sub(r'<link\s+rel="stylesheet"\s+href="([^"]+)"\s*>', css_repl, html)
    html = re.sub(r'<script\s+src="([^"]+)"\s*></script>', js_repl, html)

    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    banner = (
        f"<!-- 客户通 CustomerCRM 单文件版 · 构建于 {stamp} · "
        f"内联资源 {len(inlined)} 个：{', '.join(inlined)} -->\n"
    )
    html = html.replace("<!DOCTYPE html>", "<!DOCTYPE html>\n" + banner, 1)

    out_dir = ROOT / "dist"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / OUT_NAME
    out_path.write_text(html, encoding="utf-8", newline="\n")

    digest = hashlib.sha256(out_path.read_bytes()).hexdigest()
    size_kb = out_path.stat().st_size / 1024
    print(f"已生成：{out_path}")
    print(f"大小：{size_kb:.1f} KB")
    print(f"SHA256：{digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
