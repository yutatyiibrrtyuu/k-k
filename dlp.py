"""
DLP（数据防泄漏）加密导出工具
==============================
设计成果在本机始终保持明文，供设计人员正常编辑、读写，不受任何影响。
当需要将文件/文件夹拷出（导出给外部）时，由本方人员使用预设密码，
将其打包成加密文件（.dlpkg）。该密码只有本方人员知道并保管，
对方拿到 .dlpkg 文件后，没有本方配合（提供密码或代为解密）无法打开。

加密方案
--------
AES-256-GCM，密钥通过 PBKDF2-HMAC-SHA256（密码 + 随机 salt，20 万次迭代）派生。
打包前先用 tar+gzip 归档，再整体加密，文件内容、文件名均不可见。

预设密码管理
------------
预设密码保存在本机用户目录下的 ~/.dlp_config.json（不进版本库，已在 .gitignore 中忽略）。
只有掌握/可访问该配置文件的本方人员才能直接导出，其他人需要手动输入密码。

用法
----
设置/更新预设密码：
    python dlp.py set-password

导出加密包（命令行）：
    python dlp.py export 设计图.dwg 报告.xlsx -o 交付包.dlpkg

解密还原（命令行，需要密码）：
    python dlp.py open 交付包.dlpkg -o ./还原目录

启动简易图形界面：
    python dlp.py gui
"""

from __future__ import annotations

import argparse
import getpass
import io
import json
import os
import sys
import tarfile
from pathlib import Path
from typing import List, Optional

try:
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
except ImportError:
    sys.exit("缺少依赖：请先运行 pip install cryptography 后重试。")

MAGIC = b"DLPK1"
SALT_LEN = 16
NONCE_LEN = 12
PBKDF2_ITERATIONS = 200_000

CONFIG_PATH = Path.home() / ".dlp_config.json"


# ---------------------------------------------------------------------------
# 预设密码管理
# ---------------------------------------------------------------------------


def load_preset_password() -> Optional[str]:
    if not CONFIG_PATH.exists():
        return None
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        return data.get("password")
    except (json.JSONDecodeError, OSError):
        return None


def save_preset_password(password: str) -> None:
    CONFIG_PATH.write_text(json.dumps({"password": password}), encoding="utf-8")
    try:
        os.chmod(CONFIG_PATH, 0o600)
    except OSError:
        pass


def resolve_password(explicit: Optional[str], prompt_label: str = "密码") -> str:
    if explicit:
        return explicit
    preset = load_preset_password()
    if preset:
        return preset
    return getpass.getpass(f"请输入{prompt_label}: ")


# ---------------------------------------------------------------------------
# 加密 / 解密核心
# ---------------------------------------------------------------------------


def _derive_key(password: str, salt: bytes) -> bytes:
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=PBKDF2_ITERATIONS,
    )
    return kdf.derive(password.encode("utf-8"))


def _archive_paths(paths: List[Path]) -> bytes:
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w:gz") as tar:
        for p in paths:
            tar.add(p, arcname=p.name)
    return buf.getvalue()


def encrypt_paths(paths: List[Path], password: str, output: Path) -> None:
    archive = _archive_paths(paths)
    salt = os.urandom(SALT_LEN)
    nonce = os.urandom(NONCE_LEN)
    key = _derive_key(password, salt)
    ciphertext = AESGCM(key).encrypt(nonce, archive, None)
    with open(output, "wb") as f:
        f.write(MAGIC)
        f.write(salt)
        f.write(nonce)
        f.write(ciphertext)


def decrypt_package(package: Path, password: str, output_dir: Path) -> None:
    data = package.read_bytes()
    if data[: len(MAGIC)] != MAGIC:
        raise ValueError("不是有效的 .dlpkg 加密包")
    offset = len(MAGIC)
    salt = data[offset: offset + SALT_LEN]
    offset += SALT_LEN
    nonce = data[offset: offset + NONCE_LEN]
    offset += NONCE_LEN
    ciphertext = data[offset:]

    key = _derive_key(password, salt)
    try:
        archive = AESGCM(key).decrypt(nonce, ciphertext, None)
    except Exception as exc:
        raise ValueError("密码错误或文件已损坏，无法解密") from exc

    output_dir.mkdir(parents=True, exist_ok=True)
    with tarfile.open(fileobj=io.BytesIO(archive), mode="r:gz") as tar:
        tar.extractall(output_dir)


# ---------------------------------------------------------------------------
# 简易图形界面
# ---------------------------------------------------------------------------


def launch_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("DLP 加密导出工具")
    root.geometry("520x360")

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)

    # --- 导出 tab ---
    export_tab = ttk.Frame(notebook)
    notebook.add(export_tab, text="导出加密包")

    selected_paths: List[str] = []

    paths_listbox = tk.Listbox(export_tab, height=8)
    paths_listbox.pack(fill="both", expand=True, padx=8, pady=8)

    def add_files():
        files = filedialog.askopenfilenames(title="选择要导出的文件")
        for f in files:
            selected_paths.append(f)
            paths_listbox.insert("end", f)

    def add_folder():
        folder = filedialog.askdirectory(title="选择要导出的文件夹")
        if folder:
            selected_paths.append(folder)
            paths_listbox.insert("end", folder)

    def clear_selection():
        selected_paths.clear()
        paths_listbox.delete(0, "end")

    btn_frame = ttk.Frame(export_tab)
    btn_frame.pack(fill="x", padx=8)
    ttk.Button(btn_frame, text="添加文件", command=add_files).pack(side="left")
    ttk.Button(btn_frame, text="添加文件夹", command=add_folder).pack(side="left", padx=6)
    ttk.Button(btn_frame, text="清空", command=clear_selection).pack(side="left")

    pwd_frame = ttk.Frame(export_tab)
    pwd_frame.pack(fill="x", padx=8, pady=8)
    ttk.Label(pwd_frame, text="密码：").pack(side="left")
    export_pwd_var = tk.StringVar(value=load_preset_password() or "")
    ttk.Entry(pwd_frame, textvariable=export_pwd_var, show="*", width=30).pack(
        side="left", padx=6
    )

    def do_export():
        if not selected_paths:
            messagebox.showwarning("提示", "请先添加要导出的文件或文件夹")
            return
        password = export_pwd_var.get()
        if not password:
            messagebox.showwarning("提示", "请输入密码")
            return
        output = filedialog.asksaveasfilename(
            title="保存加密包", defaultextension=".dlpkg",
            filetypes=[("DLP 加密包", "*.dlpkg")],
        )
        if not output:
            return
        try:
            encrypt_paths([Path(p) for p in selected_paths], password, Path(output))
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        messagebox.showinfo("完成", f"已生成加密包：{output}")

    ttk.Button(export_tab, text="生成加密包", command=do_export).pack(pady=6)

    # --- 解密 tab ---
    open_tab = ttk.Frame(notebook)
    notebook.add(open_tab, text="解密导入")

    package_var = tk.StringVar()
    pkg_frame = ttk.Frame(open_tab)
    pkg_frame.pack(fill="x", padx=8, pady=8)
    ttk.Label(pkg_frame, text="加密包：").pack(side="left")
    ttk.Entry(pkg_frame, textvariable=package_var, width=30).pack(side="left", padx=6)

    def pick_package():
        f = filedialog.askopenfilename(
            title="选择 .dlpkg 加密包", filetypes=[("DLP 加密包", "*.dlpkg")]
        )
        if f:
            package_var.set(f)

    ttk.Button(pkg_frame, text="浏览", command=pick_package).pack(side="left")

    open_pwd_frame = ttk.Frame(open_tab)
    open_pwd_frame.pack(fill="x", padx=8, pady=8)
    ttk.Label(open_pwd_frame, text="密码：").pack(side="left")
    open_pwd_var = tk.StringVar()
    ttk.Entry(open_pwd_frame, textvariable=open_pwd_var, show="*", width=30).pack(
        side="left", padx=6
    )

    def do_open():
        if not package_var.get():
            messagebox.showwarning("提示", "请选择加密包文件")
            return
        if not open_pwd_var.get():
            messagebox.showwarning("提示", "请输入密码")
            return
        out_dir = filedialog.askdirectory(title="选择还原到的目录")
        if not out_dir:
            return
        try:
            decrypt_package(Path(package_var.get()), open_pwd_var.get(), Path(out_dir))
        except Exception as exc:
            messagebox.showerror("解密失败", str(exc))
            return
        messagebox.showinfo("完成", f"已还原到：{out_dir}")

    ttk.Button(open_tab, text="解密并还原", command=do_open).pack(pady=6)

    root.mainloop()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="DLP 加密导出工具")
    sub = parser.add_subparsers(dest="command", required=True)

    export_p = sub.add_parser("export", help="将文件/文件夹打包加密为 .dlpkg")
    export_p.add_argument("paths", nargs="+", type=Path, help="要导出的文件/文件夹")
    export_p.add_argument("-o", "--output", type=Path, required=True, help="输出 .dlpkg 路径")
    export_p.add_argument("--password", help="加密密码（缺省则使用预设密码或交互输入）")

    open_p = sub.add_parser("open", help="解密 .dlpkg 并还原文件")
    open_p.add_argument("package", type=Path, help=".dlpkg 加密包路径")
    open_p.add_argument("-o", "--output", type=Path, required=True, help="还原到的目录")
    open_p.add_argument("--password", help="解密密码（缺省则使用预设密码或交互输入）")

    sub.add_parser("set-password", help="设置/更新本机预设密码（保存在 ~/.dlp_config.json）")
    sub.add_parser("gui", help="启动简易图形界面")

    args = parser.parse_args(argv)

    if args.command == "set-password":
        password = getpass.getpass("请输入新的预设密码: ")
        confirm = getpass.getpass("请再次输入确认: ")
        if password != confirm:
            print("两次输入不一致，已取消。", file=sys.stderr)
            return 1
        save_preset_password(password)
        print(f"预设密码已保存到 {CONFIG_PATH}")
        return 0

    if args.command == "export":
        missing = [p for p in args.paths if not p.exists()]
        if missing:
            print(f"路径不存在：{', '.join(str(p) for p in missing)}", file=sys.stderr)
            return 1
        password = resolve_password(args.password, "加密密码")
        encrypt_paths(args.paths, password, args.output)
        print(f"已生成加密包：{args.output}")
        return 0

    if args.command == "open":
        if not args.package.is_file():
            print(f"文件不存在：{args.package}", file=sys.stderr)
            return 1
        password = resolve_password(args.password, "解密密码")
        try:
            decrypt_package(args.package, password, args.output)
        except ValueError as exc:
            print(str(exc), file=sys.stderr)
            return 1
        print(f"已还原到：{args.output}")
        return 0

    if args.command == "gui":
        launch_gui()
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
