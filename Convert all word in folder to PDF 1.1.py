import os
import sys
import re
import time
import traceback
import subprocess
from typing import List, Tuple, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def _is_windows() -> bool:
    return os.name == "nt"


def _natural_key(s: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def _ensure_dir(p: str) -> None:
    os.makedirs(p, exist_ok=True)


def _is_word_file(path: str) -> bool:
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".doc", ".docx"):
        return False
    base = os.path.basename(path)
    return not base.startswith("~$")


def _pip_install(pkgs: List[str]) -> Tuple[bool, str]:
    try:
        p = subprocess.run(
            [sys.executable, "-m", "pip", "install", *pkgs],
            capture_output=True,
            text=True,
        )
        out = (p.stdout or "") + ("\n" + p.stderr if p.stderr else "")
        return (p.returncode == 0, out.strip())
    except Exception as e:
        return (False, repr(e))


def ensure_dependencies(require_pdf_merge: bool, parent=None) -> bool:
    if not _is_windows():
        messagebox.showerror("Error", "Windows is required (Microsoft Word automation).", parent=parent)
        return False

    missing: List[str] = []

    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        import pywintypes  # noqa: F401
    except Exception:
        missing.append("pywin32")

    if require_pdf_merge:
        try:
            import pypdf  # noqa: F401
        except Exception:
            missing.append("pypdf")

    if not missing:
        return True

    if not messagebox.askyesno(
        "Install dependencies",
        "Missing dependencies detected:\n\n"
        + "\n".join(missing)
        + "\n\nInstall now into this Python environment?",
        parent=parent,
    ):
        return False

    ok, out = _pip_install(missing)
    if not ok:
        messagebox.showerror("Install failed", out or "pip install failed.", parent=parent)
        return False

    # Some environments benefit from this, but it is not always required; ignore errors.
    try:
        subprocess.run([sys.executable, "-m", "pywin32_postinstall", "-install"], capture_output=True, text=True)
    except Exception:
        pass

    # Re-check imports
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        import pywintypes  # noqa: F401
    except Exception as e:
        messagebox.showerror("Dependency error", f"pywin32 still not usable:\n\n{repr(e)}", parent=parent)
        return False

    if require_pdf_merge:
        try:
            import pypdf  # noqa: F401
        except Exception as e:
            messagebox.showerror("Dependency error", f"pypdf still not usable:\n\n{repr(e)}", parent=parent)
            return False

    return True


# Word constants (numeric to avoid dependency on generated constants)
WD_FORMAT_PDF = 17
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def _create_word_app():
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        word.DisplayAlerts = 0
    except Exception:
        pass
    try:
        word.Options.SaveNormalPrompt = False
    except Exception:
        pass
    try:
        word.Options.ConfirmConversions = False
    except Exception:
        pass
    try:
        word.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception:
        pass
    return word


def _open_doc_hardened(word, path: str):
    ap = os.path.abspath(path)

    # Attempt normal open first
    try:
        return word.Documents.Open(
            ap,
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False,
            NoEncodingDialog=True,
            Revert=False,
            OpenAndRepair=True,
            Visible=False,
        )
    except Exception:
        pass

    # Protected View fallback
    pv = word.ProtectedViewWindows.Open(ap, AddToRecentFiles=False)
    return pv.Edit()


def _save_as_pdf(doc, pdf_path: str):
    out = os.path.abspath(pdf_path)
    _ensure_dir(os.path.dirname(out))
    # SaveAs2 tends to be more reliable than ExportAsFixedFormat in automation-heavy contexts
    doc.SaveAs2(out, FileFormat=WD_FORMAT_PDF, AddToRecentFiles=False)


def convert_folder(
    input_root: str,
    output_root: str,
    overwrite: bool,
    progress_cb=None,
    log_path: Optional[str] = None,
) -> Tuple[List[str], List[Tuple[str, str]]]:
    import pythoncom

    word_files: List[str] = []
    for dirpath, _, filenames in os.walk(input_root):
        for fn in filenames:
            fp = os.path.join(dirpath, fn)
            if _is_word_file(fp):
                word_files.append(fp)
    word_files.sort(key=_natural_key)

    pdfs: List[str] = []
    failures: List[Tuple[str, str]] = []

    pythoncom.CoInitialize()
    try:
        word = _create_word_app()
        try:
            for i, src in enumerate(word_files, start=1):
                try:
                    rel = os.path.relpath(src, input_root)
                    rel_dir = os.path.dirname(rel)
                    base = os.path.splitext(os.path.basename(src))[0] + ".pdf"
                    out_dir = os.path.join(output_root, rel_dir) if rel_dir else output_root
                    out_pdf = os.path.join(out_dir, base)

                    if os.path.exists(out_pdf) and not overwrite:
                        pdfs.append(out_pdf)
                    else:
                        doc = None
                        try:
                            doc = _open_doc_hardened(word, src)
                            _save_as_pdf(doc, out_pdf)
                            pdfs.append(out_pdf)
                        finally:
                            if doc is not None:
                                try:
                                    doc.Close(SaveChanges=False)
                                except Exception:
                                    pass

                except Exception as e:
                    failures.append((src, repr(e)))
                    if log_path:
                        with open(log_path, "a", encoding="utf-8") as f:
                            f.write(f"FAIL: {src} | {repr(e)}\n")

                if progress_cb:
                    progress_cb(i, len(word_files))

        finally:
            try:
                word.Quit()
            except Exception:
                pass

    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return pdfs, failures


def merge_pdfs(pdf_paths: List[str], merged_output: str):
    from pypdf import PdfReader, PdfWriter

    writer = PdfWriter()
    for p in sorted(pdf_paths, key=_natural_key):
        if not os.path.exists(p):
            continue
        r = PdfReader(p)
        for page in r.pages:
            writer.add_page(page)

    _ensure_dir(os.path.dirname(os.path.abspath(merged_output)))
    with open(merged_output, "wb") as f:
        writer.write(f)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DOC/DOCX to PDF (Word)")
        self.geometry("740x300")
        self.resizable(False, False)

        self.in_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.overwrite_var = tk.BooleanVar(value=False)
        self.merge_var = tk.BooleanVar(value=False)

        self._build()

    def _build(self):
        pad = 10
        frm = ttk.Frame(self, padding=pad)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Input folder").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.in_var, width=75).grid(row=1, column=0, sticky="w")
        ttk.Button(frm, text="Select...", command=self.pick_input).grid(row=1, column=1, sticky="w")

        ttk.Label(frm, text="Output folder").grid(row=2, column=0, sticky="w", pady=(pad, 0))
        ttk.Entry(frm, textvariable=self.out_var, width=75).grid(row=3, column=0, sticky="w")
        ttk.Button(frm, text="Select...", command=self.pick_output).grid(row=3, column=1, sticky="w")

        ttk.Checkbutton(frm, text="Overwrite existing PDFs", variable=self.overwrite_var).grid(row=4, column=0, sticky="w", pady=(pad, 0))
        ttk.Checkbutton(frm, text="Merge all PDFs into one", variable=self.merge_var).grid(row=5, column=0, sticky="w")

        self.prog = ttk.Progressbar(frm, length=620, mode="determinate")
        self.prog.grid(row=6, column=0, sticky="w", pady=(pad, 0))
        self.status = ttk.Label(frm, text="Idle.")
        self.status.grid(row=7, column=0, sticky="w")

        ttk.Button(frm, text="Run", command=self.run).grid(row=8, column=0, sticky="w", pady=(pad, 0))

        frm.grid_columnconfigure(0, weight=1)

    def pick_input(self):
        d = filedialog.askdirectory(title="Select input folder")
        if d:
            self.in_var.set(d)
            self.out_var.set(os.path.join(d, "PDF_Output"))

    def pick_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self.out_var.set(d)

    def _set_progress(self, cur: int, total: int):
        self.prog["maximum"] = max(total, 1)
        self.prog["value"] = cur
        self.status.config(text=f"Processed {cur} of {total}...")
        self.update_idletasks()

    def run(self):
        inp = self.in_var.get().strip()
        outp = self.out_var.get().strip()
        do_merge = bool(self.merge_var.get())

        if not inp or not os.path.isdir(inp):
            messagebox.showerror("Error", "Select a valid input folder.", parent=self)
            return
        if not outp:
            messagebox.showerror("Error", "Select a valid output folder.", parent=self)
            return

        if not ensure_dependencies(require_pdf_merge=do_merge, parent=self):
            return

        _ensure_dir(outp)
        log_path = os.path.join(outp, "conversion_log.txt")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"Run started: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")

        self.prog["value"] = 0
        self.status.config(text="Starting...")
        self.update_idletasks()

        pdfs, failures = convert_folder(
            input_root=inp,
            output_root=outp,
            overwrite=bool(self.overwrite_var.get()),
            progress_cb=self._set_progress,
            log_path=log_path,
        )

        if do_merge and pdfs:
            merged = os.path.join(outp, "Merged_All.pdf")
            try:
                merge_pdfs(pdfs, merged)
            except Exception as e:
                with open(log_path, "a", encoding="utf-8") as f:
                    f.write(f"MERGE FAIL: {repr(e)}\n")
                messagebox.showwarning("Merge failed", f"Merge failed:\n\n{repr(e)}\n\nSee log:\n{log_path}", parent=self)

        if failures:
            messagebox.showwarning(
                "Done (with failures)",
                f"Converted: {len(pdfs)}\nFailures: {len(failures)}\n\nSee log:\n{log_path}",
                parent=self,
            )
        else:
            messagebox.showinfo(
                "Done",
                f"Converted: {len(pdfs)}\n\nSee log:\n{log_path}",
                parent=self,
            )

        self.status.config(text="Finished.")


def main():
    if not _is_windows():
        print("Windows is required (Microsoft Word automation).")
        return
    App().mainloop()


if __name__ == "__main__":
    main()
