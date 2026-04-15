"""
debug_run_catvbs.py
===================
诊断脚本：逐一尝试多种方法让 CATIA 通过 COM 执行 .catvbs 文件。

用法（在 CATIA 已启动的 Windows 机器上）：
    python debug_run_catvbs.py

脚本会自动使用 macros/ 目录中的 generate_drawing.catvbs，
也可以在命令行指定其他 .catvbs 文件：
    python debug_run_catvbs.py "D:\\path\\to\\your_macro.catvbs"

每种方法的结果（成功 / 失败 + 异常信息）会打印到控制台，
并在末尾生成一份 debug_run_catvbs_result.txt 报告。
"""

import sys
import os
import traceback
from pathlib import Path

# ---------------------------------------------------------------------------
# 0. 确定 .catvbs 文件路径
# ---------------------------------------------------------------------------
_HERE = Path(__file__).resolve().parent
MACRO_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else (
    _HERE / "macros" / "generate_drawing.catvbs"
)

REPORT_FILE = _HERE / "debug_run_catvbs_result.txt"

results = []  # list of (method_name, success: bool, detail: str)


def log(method: str, ok: bool, detail: str = "") -> None:
    tag = "✅ 成功" if ok else "❌ 失败"
    line = f"[{tag}] {method}"
    if detail:
        line += f"\n        {detail}"
    print(line)
    results.append((method, ok, detail))


# ---------------------------------------------------------------------------
# 1. 检查文件是否存在
# ---------------------------------------------------------------------------
print("=" * 60)
print(f"调试目标文件: {MACRO_PATH}")
print("=" * 60)

if not MACRO_PATH.exists():
    print(f"❌ 文件不存在: {MACRO_PATH}")
    sys.exit(1)

LIB_DIR = str(MACRO_PATH.parent)
MACRO_NAME = MACRO_PATH.name
MACRO_FULL = str(MACRO_PATH)

# ---------------------------------------------------------------------------
# 2. 连接 CATIA COM 对象
# ---------------------------------------------------------------------------
print("\n--- 连接 CATIA ---")

app = None

# 方法 A：win32com.client.GetActiveObject (最推荐)
try:
    import win32com.client
    app = win32com.client.GetActiveObject("CATIA.Application")
    print(f"✅ win32com GetActiveObject('CATIA.Application') 连接成功")
except Exception as e:
    print(f"⚠  win32com GetActiveObject('CATIA.Application') 失败: {e}")

if app is None:
    # 方法 B：尝试其他常见 ProgID（V5R21 / V5-6R2014 等）
    for progid in ["CATIA.Application", "CATIAApplicationClass",
                   "CNEXT.Application", "CATIAApplication"]:
        try:
            import win32com.client
            app = win32com.client.GetActiveObject(progid)
            print(f"✅ win32com GetActiveObject('{progid}') 连接成功")
            break
        except Exception as e:
            print(f"   GetActiveObject('{progid}'): {e}")

if app is None:
    # 方法 C：pycatia
    try:
        from pycatia import catia as _catia
        caa = _catia()
        app = caa.application.com_object
        print("✅ pycatia catia().application.com_object 连接成功")
    except Exception as e:
        print(f"⚠  pycatia 连接失败: {e}")

if app is None:
    print("\n❌ 所有方法均无法连接 CATIA，请确保 CATIA 已启动后重试。")
    sys.exit(1)

print(f"\n当前活动文档: ", end="")
try:
    doc_name = app.ActiveDocument.Name
    doc_type_str = app.SystemService.Evaluate(
        "TypeName(CATIA.ActiveDocument)", 0, "", []
    )
    print(f"{doc_name}  (类型: {doc_type_str})")
except Exception as e:
    print(f"(无法读取活动文档: {e})")

sys_svc = app.SystemService
print(f"\n>>> 宏目录  : {LIB_DIR}")
print(f">>> 宏文件名: {MACRO_NAME}")
print(f">>> 完整路径: {MACRO_FULL}")

# ---------------------------------------------------------------------------
# 3. 逐一尝试各种 ExecuteScript 调用方式
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("开始逐一尝试方法")
print("=" * 60)

# ── 方法 1：LibraryType=0，先不注册，直接传目录 + 文件名 ──────────────────
try:
    sys_svc.ExecuteScript(LIB_DIR, 0, MACRO_NAME, "CATMain", [])
    log("方法1: ExecuteScript(dir, 0, filename) — 不注册宏库", True)
except Exception as e:
    log("方法1: ExecuteScript(dir, 0, filename) — 不注册宏库", False,
        str(e).split("\n")[0])

# ── 方法 2：先 MacroLibraries.Add 注册目录，再 ExecuteScript ──────────────
try:
    try:
        sys_svc.MacroLibraries.Add(LIB_DIR, 0)
        print(f"   MacroLibraries.Add({LIB_DIR!r}, 0) 成功")
    except Exception as add_err:
        print(f"   MacroLibraries.Add 异常（继续尝试）: {add_err}")
    sys_svc.ExecuteScript(LIB_DIR, 0, MACRO_NAME, "CATMain", [])
    log("方法2: MacroLibraries.Add + ExecuteScript(dir, 0, filename)", True)
except Exception as e:
    log("方法2: MacroLibraries.Add + ExecuteScript(dir, 0, filename)", False,
        str(e).split("\n")[0])

# ── 方法 3：LibraryType=0，传完整路径作为 iLibraryName ──────────────────
try:
    sys_svc.ExecuteScript(MACRO_FULL, 0, MACRO_NAME, "CATMain", [])
    log("方法3: ExecuteScript(full_path, 0, filename)", True)
except Exception as e:
    log("方法3: ExecuteScript(full_path, 0, filename)", False,
        str(e).split("\n")[0])

# ── 方法 4：LibraryType=0，iProgramName 传完整路径 ────────────────────────
try:
    sys_svc.ExecuteScript(LIB_DIR, 0, MACRO_FULL, "CATMain", [])
    log("方法4: ExecuteScript(dir, 0, full_path)", True)
except Exception as e:
    log("方法4: ExecuteScript(dir, 0, full_path)", False,
        str(e).split("\n")[0])

# ── 方法 5：LibraryType=0，两个参数都传完整路径 ───────────────────────────
try:
    sys_svc.ExecuteScript(MACRO_FULL, 0, MACRO_FULL, "CATMain", [])
    log("方法5: ExecuteScript(full_path, 0, full_path)", True)
except Exception as e:
    log("方法5: ExecuteScript(full_path, 0, full_path)", False,
        str(e).split("\n")[0])

# ── 方法 6：LibraryType=0，iProgramName 传不含扩展名的文件名 ─────────────
MACRO_STEM = MACRO_PATH.stem  # e.g. "generate_drawing"
try:
    sys_svc.ExecuteScript(LIB_DIR, 0, MACRO_STEM, "CATMain", [])
    log(f"方法6: ExecuteScript(dir, 0, '{MACRO_STEM}') — 无扩展名", True)
except Exception as e:
    log(f"方法6: ExecuteScript(dir, 0, '{MACRO_STEM}') — 无扩展名", False,
        str(e).split("\n")[0])

# ── 方法 7：LibraryType=0，用反斜杠标准化目录路径 ─────────────────────────
LIB_DIR_BS = str(MACRO_PATH.parent).replace("/", "\\").rstrip("\\")
try:
    sys_svc.ExecuteScript(LIB_DIR_BS, 0, MACRO_NAME, "CATMain", [])
    log(f"方法7: ExecuteScript(反斜杠路径, 0, filename)", True)
except Exception as e:
    log(f"方法7: ExecuteScript(反斜杠路径, 0, filename)", False,
        str(e).split("\n")[0])

# ── 方法 8：LibraryType=0，目录末尾带反斜杠 ──────────────────────────────
LIB_DIR_TRAIL = LIB_DIR_BS + "\\"
try:
    sys_svc.ExecuteScript(LIB_DIR_TRAIL, 0, MACRO_NAME, "CATMain", [])
    log(f"方法8: ExecuteScript(目录\\, 0, filename)", True)
except Exception as e:
    log(f"方法8: ExecuteScript(目录\\, 0, filename)", False,
        str(e).split("\n")[0])

# ── 方法 9：先用 Evaluate 执行 "AddMacroLibrary" 语句，再 ExecuteScript ──
try:
    add_lib_vbs = f'CATIA.SystemService.MacroLibraries.Add "{LIB_DIR_BS}", 0'
    sys_svc.Evaluate(add_lib_vbs, 0, "", [])
    sys_svc.ExecuteScript(LIB_DIR, 0, MACRO_NAME, "CATMain", [])
    log("方法9: Evaluate 添加宏库 + ExecuteScript", True)
except Exception as e:
    log("方法9: Evaluate 添加宏库 + ExecuteScript", False,
        str(e).split("\n")[0])

# ── 方法 10：用 Evaluate 直接执行整个脚本内容（inline eval）──────────────
try:
    script_content = MACRO_PATH.read_text(encoding="utf-8-sig", errors="replace")
    sys_svc.Evaluate(script_content, 0, "CATMain", [])
    log("方法10: Evaluate(脚本内容, 0, 'CATMain', [])", True)
except Exception as e:
    log("方法10: Evaluate(脚本内容, 0, 'CATMain', [])", False,
        str(e).split("\n")[0])

# ── 方法 11：用 Evaluate，第二参数传 1 ───────────────────────────────────
try:
    script_content = MACRO_PATH.read_text(encoding="utf-8-sig", errors="replace")
    sys_svc.Evaluate(script_content, 1, "CATMain", [])
    log("方法11: Evaluate(脚本内容, 1, 'CATMain', [])", True)
except Exception as e:
    log("方法11: Evaluate(脚本内容, 1, 'CATMain', [])", False,
        str(e).split("\n")[0])

# ── 方法 12：LibraryType=1 (VBA)，传完整路径 ─────────────────────────────
try:
    sys_svc.ExecuteScript(MACRO_FULL, 1, "Module1", "CATMain", [])
    log("方法12: ExecuteScript(full_path, 1=VBA, 'Module1', ...)", True)
except Exception as e:
    log("方法12: ExecuteScript(full_path, 1=VBA, 'Module1', ...)", False,
        str(e).split("\n")[0])

# ── 方法 13：枚举当前宏库列表（诊断信息） ────────────────────────────────
print("\n--- 诊断：列出当前已注册的宏库 ---")
try:
    libs = sys_svc.MacroLibraries
    count = libs.Count
    print(f"已注册宏库数量: {count}")
    for i in range(1, count + 1):
        lib = libs.Item(i)
        try:
            print(f"  [{i}] Path={lib.Path}  Type={lib.Type}")
        except Exception:
            print(f"  [{i}] (无法读取属性)")
    log(f"方法13(诊断): 枚举 MacroLibraries ({count} 个)", True,
        f"已注册库数量: {count}")
except Exception as e:
    log("方法13(诊断): MacroLibraries 枚举", False, str(e).split("\n")[0])

# ---------------------------------------------------------------------------
# 4. 汇总结果
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("汇总结果")
print("=" * 60)

success_methods = [m for m, ok, _ in results if ok]
fail_methods    = [m for m, ok, _ in results if not ok]

print(f"✅ 成功: {len(success_methods)} 种方法")
for m in success_methods:
    print(f"   • {m}")

print(f"\n❌ 失败: {len(fail_methods)} 种方法")
for m, ok, detail in results:
    if not ok:
        print(f"   • {m}")
        if detail:
            print(f"       → {detail}")

# 写入报告文件
with open(REPORT_FILE, "w", encoding="utf-8") as f:
    f.write(f"CATIA .catvbs 执行方法诊断报告\n")
    f.write(f"目标文件: {MACRO_FULL}\n")
    f.write("=" * 60 + "\n\n")
    for method, ok, detail in results:
        tag = "✅ 成功" if ok else "❌ 失败"
        f.write(f"[{tag}] {method}\n")
        if detail:
            f.write(f"        {detail}\n")
        f.write("\n")

print(f"\n报告已保存至: {REPORT_FILE}")
