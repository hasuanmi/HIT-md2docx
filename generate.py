#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate.py — HIT 学位论文「生成 Word → 注入官方封面」一条龙

用法：
  python generate.py <input.md> [output.docx] [--profile hit-master-thesis]
                  [--no-cover] [--degree master|bachelor|doctor]

等价于依次执行：
  1. md2docx docx --profile <profile> <input.md> <temp.docx>
  2. skill/scripts/cover_inject.py <temp.docx> --cover-template input/封面.docx
     （--no-cover 时跳过）
  3. degree_adapt（非 master 时自动执行）
  4. 输出到 <output.docx> 或 <input同名>.docx

找不到 封面.docx 时自动跳过封面注入并打印提示。
"""

import argparse
import os
import shutil
import subprocess
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ENGINE_ROOT = SCRIPT_DIR


def main():
    parser = argparse.ArgumentParser(description="HIT 论文: md → docx + 官方封面")
    parser.add_argument("input_md", help="输入 Markdown 文件")
    parser.add_argument("output_docx", nargs="?", help="输出 .docx 文件（默认: <输入同名>.docx）")
    parser.add_argument("--profile", default="hit-master-thesis", help="引擎 profile（默认 hit-master-thesis）")
    parser.add_argument("--no-cover", action="store_true", help="跳过官方封面模板注入")
    parser.add_argument("--degree", default="master",
                        choices=["master", "bachelor", "doctor"],
                        help="学位层次（默认 master，仅 bachelor/doctor 触发改写）")
    args = parser.parse_args()

    input_md = os.path.abspath(args.input_md)
    if not os.path.isfile(input_md):
        print(f"错误：文件不存在 {input_md}", file=sys.stderr)
        sys.exit(1)

    # 确定输出路径
    base = os.path.splitext(os.path.basename(input_md))[0]
    output_docx = args.output_docx or f"{base}.docx"
    output_docx = os.path.abspath(output_docx)
    temp_docx = output_docx + ".tmp"

    py = sys.executable

    # ---- Step 1: 引擎生成 ----
    # 注意：必须走 run_engine.py（它会把【当前仓库】的 thesis_md2docx 用
    # sys.path.insert(0, REPO) 顶到最前面），避免本机存在旧版/可编辑安装的
    # thesis_md2docx 被 Python 优先导入（"引擎 shadow" 坑，会导致封面错成
    # 本科学位论文等旧行为）。直接 `python -m thesis_md2docx.main` 在旧安装
    # 存在时会命中 shadow，故这里默认走 run_engine.py，缺失时再回退。
    print(f"[1/3] 引擎生成: {input_md} -> {os.path.basename(temp_docx)} "
          f"(profile={args.profile})")
    run_engine = os.path.join(ENGINE_ROOT, "skill", "scripts", "run_engine.py")
    if os.path.isfile(run_engine):
        cmd1 = [py, run_engine, "docx",
                "--profile", args.profile, input_md, temp_docx]
    else:
        cmd1 = [py, "-m", "thesis_md2docx.main", "docx",
                "--profile", args.profile, input_md, temp_docx]
    r = subprocess.run(cmd1, cwd=ENGINE_ROOT)
    if r.returncode != 0:
        print(f"引擎失败 (exit {r.returncode})", file=sys.stderr)
        sys.exit(r.returncode)

    # ---- Step 2: 注入官方封面 ----
    cover_template = os.path.join(ENGINE_ROOT, "input", "封面.docx")
    cover_script = os.path.join(ENGINE_ROOT, "skill", "scripts", "cover_inject.py")
    cover_ok = False  # 追踪注入是否真正成功

    if args.no_cover:
        print("[2/3] 封面注入: 跳过 (--no-cover)")
    elif not os.path.isfile(cover_script):
        print(f"[2/3] 封面注入: 跳过（脚本不存在 {cover_script}）")
    elif not os.path.isfile(cover_template):
        print(f"[2/3] ⚠ 官方封面模板未找到 ({cover_template})，"
              f"将使用引擎内置封面。如需官方封面版式，请把学校下发的 封面.docx "
              f"放到 input/ 目录下。")
    else:
        print(f"[2/3] 注入官方封面模板...")
        cmd2 = [
            py, cover_script,
            temp_docx,
            "--cover-template", cover_template,
            "--out", temp_docx,
        ]
        r2 = subprocess.run(cmd2, cwd=ENGINE_ROOT,
                            capture_output=True, text=True)
        print(r2.stdout)
        if r2.stderr:
            print(r2.stderr, file=sys.stderr)
        if r2.returncode != 0:
            print(f"  ⚠ 封面注入返回非零 ({r2.returncode})，保留引擎内置封面")
        elif "封面区注入: 成功" in r2.stdout or "封面区注入: 成功" in r2.stderr:
            cover_ok = True
        elif "跳过" in r2.stdout or "跳过" in r2.stderr:
            print("  ❌ 封面注入被跳过！输出可能使用引擎内置封面（版式不正确）。"
                  "请检查上方的警告信息。")

    # ---- Step 2.5: 封面注入验证（用 lxml 快速扫 document.xml）----
    if not args.no_cover and os.path.isfile(temp_docx):
        try:
            import zipfile
            from lxml import etree
            with zipfile.ZipFile(temp_docx, "r") as zf:
                xml_bytes = zf.read("word/document.xml")
            root = etree.fromstring(xml_bytes)
            W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            all_text = "".join(root.iter(f"{{{W}}}t")).itertext()
            # 官方封面模板的标志性文字（来自 input/封面.docx）
            OFFICIAL_MARKERS = ["哈尔滨工业大学", "硕士学位论文"]
            found = [m for m in OFFICIAL_MARKERS if m in all_text]
            if len(found) >= 2:
                if not cover_ok:
                    cover_ok = True  # lxml 验证确认成功
                print(f"  ✓ 封面验证通过：检测到官方模板标记 {found}")
            else:
                print(f"  ⚠ 封面验证异常：文档中未检测到足够的官方封面标记 "
                      f"（找到: {found}，期望: {OFFICIAL_MARKERS}）。"
                      f"输出可能使用了引擎内置封面（布局/文字可能与学校模板不符）。"
                      f"建议检查 input/封面.docx 是否存在且未被损坏。")
        except Exception as ex:
            print(f"  ℹ 封面验证跳过（解析错误: {ex}）")

    # ---- Step 3: 学位层次适配 ----
    if args.degree in ("bachelor", "doctor"):
        degree_script = os.path.join(ENGINE_ROOT, "skill", "scripts", "degree_adapt.py")
        print(f"[3/3] 层次适配: {args.degree}")
        cmd3 = [py, degree_script, "--degree", args.degree, temp_docx]
        r3 = subprocess.run(cmd3, cwd=ENGINE_ROOT)
        if r3.returncode != 0:
            print(f"  ⚠ 层次适配返回非零 ({r3.returncode})")
    else:
        print("[3/3] 层次适配: master（无需改写）")

    # ---- 最终输出 ----
    shutil.move(temp_docx, output_docx)
    print(f"\n完成! 输出: {output_docx}")
    if not args.no_cover and not cover_ok:
        print("\n⚠️  警告：官方封面模板可能未成功注入！")
        print("  封面可能显示引擎内置版式（非学校正式模板）。")
        print("  请检查：")
        print("    1. input/封面.docx 是否存在且未被 git LFS 截断")
        print("    2. 上方 [2/3] 步骤是否有警告信息")


if __name__ == "__main__":
    main()
