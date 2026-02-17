# app_gui.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from compare_core import (
    is_excel,
    get_excel_sheet_names,
    auto_detect_header_and_load,
    index_to_excel_col_letter,
    auto_pick_best_column_index,
    compare_files,
    xlookup_join,
    differences_report,
)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("مقایسه Excel/CSV + XLOOKUP + Differences (GUI)")
        self.geometry("1100x780")

        self.file_a = tk.StringVar()
        self.file_b = tk.StringVar()
        self.sheet_a = tk.StringVar()
        self.sheet_b = tk.StringVar()

        self.col_a = tk.StringVar()
        self.col_b = tk.StringVar()
        self.out_path = tk.StringVar(value="result.xlsx")

        self.case_insensitive = tk.BooleanVar(value=False)
        self.keep_duplicates = tk.BooleanVar(value=False)  # compare only
        self.keep_blanks = tk.BooleanVar(value=False)

        self.pick_mode = tk.StringVar(value="auto")   # auto | manual
        self.action = tk.StringVar(value="compare")   # compare | lookup | diff

        self.df_a = None
        self.df_b = None
        self.header_note_a = ""
        self.header_note_b = ""

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        frm_files = ttk.LabelFrame(root, text="فایل‌ها", padding=10)
        frm_files.pack(fill="x")

        ttk.Label(frm_files, text="فایل A:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_files, textvariable=self.file_a, width=86).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm_files, text="انتخاب...", command=self.pick_a).grid(row=0, column=2)

        ttk.Label(frm_files, text="فایل B:").grid(row=1, column=0, sticky="w", pady=(8,0))
        ttk.Entry(frm_files, textvariable=self.file_b, width=86).grid(row=1, column=1, sticky="we", padx=6, pady=(8,0))
        ttk.Button(frm_files, text="انتخاب...", command=self.pick_b).grid(row=1, column=2, pady=(8,0))

        frm_files.columnconfigure(1, weight=1)

        frm_keys = ttk.LabelFrame(root, text="شیت و ستون کلید", padding=10)
        frm_keys.pack(fill="x", pady=10)

        ttk.Label(frm_keys, text="Sheet A:").grid(row=0, column=0, sticky="w")
        self.sheet_a_cb = ttk.Combobox(frm_keys, textvariable=self.sheet_a, width=30, state="readonly")
        self.sheet_a_cb.grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(frm_keys, text="Sheet B:").grid(row=0, column=2, sticky="w")
        self.sheet_b_cb = ttk.Combobox(frm_keys, textvariable=self.sheet_b, width=30, state="readonly")
        self.sheet_b_cb.grid(row=0, column=3, sticky="w", padx=6)

        ttk.Radiobutton(frm_keys, text="Auto (پیشنهادی)", variable=self.pick_mode, value="auto", command=self.refresh_columns).grid(row=1, column=0, sticky="w", pady=(10,0))
        ttk.Radiobutton(frm_keys, text="Manual (انتخاب دستی)", variable=self.pick_mode, value="manual", command=self.refresh_columns).grid(row=1, column=1, sticky="w", pady=(10,0))

        ttk.Label(frm_keys, text="Key Column A:").grid(row=2, column=0, sticky="w", pady=(10,0))
        self.col_a_cb = ttk.Combobox(frm_keys, textvariable=self.col_a, width=30, state="readonly")
        self.col_a_cb.grid(row=2, column=1, sticky="w", padx=6, pady=(10,0))

        ttk.Label(frm_keys, text="Key Column B:").grid(row=2, column=2, sticky="w", pady=(10,0))
        self.col_b_cb = ttk.Combobox(frm_keys, textvariable=self.col_b, width=30, state="readonly")
        self.col_b_cb.grid(row=2, column=3, sticky="w", padx=6, pady=(10,0))

        ttk.Button(frm_keys, text="Load / Preview", command=self.load_and_preview).grid(row=0, column=4, padx=(12,0))
        ttk.Button(frm_keys, text="Refresh Columns", command=self.refresh_columns).grid(row=2, column=4, padx=(12,0), pady=(10,0))

        frm_action = ttk.LabelFrame(root, text="نوع عملیات", padding=10)
        frm_action.pack(fill="x")

        ttk.Radiobutton(frm_action, text="Compare (مشترک/فقط در A/B)", variable=self.action, value="compare").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(frm_action, text="Lookup (XLOOKUP) - آوردن ستون‌های B داخل A", variable=self.action, value="lookup").grid(row=0, column=1, sticky="w", padx=12)
        ttk.Radiobutton(frm_action, text="Differences (مغایرت ستون‌ها برای کلیدهای مشترک)", variable=self.action, value="diff").grid(row=0, column=2, sticky="w", padx=12)

        ttk.Label(frm_action, text="ستون‌هایی از B که به A اضافه شوند (Lookup):").grid(row=1, column=0, sticky="w", pady=(10,0))
        self.bcols_list = tk.Listbox(frm_action, selectmode="extended", height=5, exportselection=False)
        self.bcols_list.grid(row=2, column=0, columnspan=2, sticky="we", pady=(6,0))

        ttk.Label(frm_action, text="ستون‌های مشترک برای مقایسه (Differences):").grid(row=1, column=2, sticky="w", pady=(10,0))
        self.diffcols_list = tk.Listbox(frm_action, selectmode="extended", height=5, exportselection=False)
        self.diffcols_list.grid(row=2, column=2, sticky="we", pady=(6,0))

        frm_action.columnconfigure(1, weight=1)
        frm_action.columnconfigure(2, weight=1)

        frm_opts = ttk.LabelFrame(root, text="تنظیمات", padding=10)
        frm_opts.pack(fill="x", pady=10)

        ttk.Checkbutton(frm_opts, text="حساس نبودن به حروف بزرگ/کوچک", variable=self.case_insensitive).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(frm_opts, text="نگه داشتن تکراری‌ها (Occurrences) [Compare]", variable=self.keep_duplicates).grid(row=0, column=1, sticky="w", padx=12)
        ttk.Checkbutton(frm_opts, text="نگه داشتن خالی‌ها", variable=self.keep_blanks).grid(row=0, column=2, sticky="w", padx=12)

        ttk.Label(frm_opts, text="خروجی (xlsx):").grid(row=1, column=0, sticky="w", pady=(10,0))
        ttk.Entry(frm_opts, textvariable=self.out_path, width=76).grid(row=1, column=1, sticky="w", pady=(10,0))
        ttk.Button(frm_opts, text="Save As...", command=self.pick_out).grid(row=1, column=2, sticky="w", pady=(10,0))

        frm_run = ttk.Frame(root)
        frm_run.pack(fill="x")

        ttk.Button(frm_run, text="Run", command=self.run_action).pack(side="left")
        self.status = ttk.Label(frm_run, text="آماده", anchor="w")
        self.status.pack(side="left", padx=12)

        frm_prev = ttk.LabelFrame(root, text="پیش‌نمایش داده", padding=10)
        frm_prev.pack(fill="both", expand=True, pady=10)

        self.txt = tk.Text(frm_prev, height=16, wrap="none")
        self.txt.pack(fill="both", expand=True)

    def pick_a(self):
        p = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")])
        if p:
            self.file_a.set(p)
            self.fill_sheets()
            self.load_and_preview()

    def pick_b(self):
        p = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.xls *.csv"), ("All", "*.*")])
        if p:
            self.file_b.set(p)
            self.fill_sheets()
            self.load_and_preview()

    def pick_out(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.out_path.set(p)

    def fill_sheets(self):
        for which in ("a", "b"):
            path = self.file_a.get() if which == "a" else self.file_b.get()
            cb = self.sheet_a_cb if which == "a" else self.sheet_b_cb
            var = self.sheet_a if which == "a" else self.sheet_b

            if path and is_excel(path):
                try:
                    names = get_excel_sheet_names(path)
                    cb["values"] = names
                    var.set(names[0] if names else "")
                except Exception:
                    cb["values"] = []
                    var.set("")
            else:
                cb["values"] = []
                var.set("")

    def load_and_preview(self):
        try:
            if self.file_a.get():
                df, _, mode = auto_detect_header_and_load(self.file_a.get(), self.sheet_a.get() or None)
                self.df_a = df
                self.header_note_a = mode
            if self.file_b.get():
                df, _, mode = auto_detect_header_and_load(self.file_b.get(), self.sheet_b.get() or None)
                self.df_b = df
                self.header_note_b = mode

            self.refresh_columns()
            self.preview()
        except Exception as e:
            messagebox.showerror("خطا", str(e))

    def refresh_columns(self):
        def build_col_list(df):
            if df is None:
                return []
            items = []
            for i, c in enumerate(list(df.columns)):
                letter = index_to_excel_col_letter(i)
                items.append(f"{letter} | {c}")
            return items

        a_list = build_col_list(self.df_a)
        b_list = build_col_list(self.df_b)
        self.col_a_cb["values"] = a_list
        self.col_b_cb["values"] = b_list

        if self.pick_mode.get() == "auto":
            if self.df_a is not None and a_list:
                idx = auto_pick_best_column_index(self.df_a)
                self.col_a.set(a_list[idx])
            if self.df_b is not None and b_list:
                idx = auto_pick_best_column_index(self.df_b)
                self.col_b.set(b_list[idx])
        else:
            if not self.col_a.get() and a_list:
                self.col_a.set(a_list[0])
            if not self.col_b.get() and b_list:
                self.col_b.set(b_list[0])

        self.bcols_list.delete(0, "end")
        if self.df_b is not None:
            for c in list(self.df_b.columns):
                self.bcols_list.insert("end", str(c))

        self.diffcols_list.delete(0, "end")
        if self.df_a is not None and self.df_b is not None:
            common = [c for c in self.df_a.columns if c in self.df_b.columns]
            for c in common:
                self.diffcols_list.insert("end", str(c))

    def _extract_letter(self, combo_value: str):
        if not combo_value:
            return None
        return combo_value.split("|", 1)[0].strip()

    def preview(self):
        self.txt.delete("1.0", "end")
        if self.df_a is not None:
            self.txt.insert("end", f"=== A ({self.header_note_a}) shape={self.df_a.shape} ===\n")
            self.txt.insert("end", f"columns: {list(self.df_a.columns)}\n")
            self.txt.insert("end", self.df_a.head(12).to_string(index=False))
            self.txt.insert("end", "\n\n")
        if self.df_b is not None:
            self.txt.insert("end", f"=== B ({self.header_note_b}) shape={self.df_b.shape} ===\n")
            self.txt.insert("end", f"columns: {list(self.df_b.columns)}\n")
            self.txt.insert("end", self.df_b.head(12).to_string(index=False))
            self.txt.insert("end", "\n")

    def run_action(self):
        if not self.file_a.get() or not self.file_b.get():
            messagebox.showwarning("هشدار", "لطفاً هر دو فایل را انتخاب کن.")
            return

        try:
            key_a = self._extract_letter(self.col_a.get())
            key_b = self._extract_letter(self.col_b.get())

            self.status.config(text="در حال اجرا...")
            self.update_idletasks()

            if self.action.get() == "compare":
                res = compare_files(
                    file_a=self.file_a.get(),
                    file_b=self.file_b.get(),
                    sheet_a=self.sheet_a.get() or None,
                    sheet_b=self.sheet_b.get() or None,
                    col_a=key_a if self.pick_mode.get() == "manual" else None,
                    col_b=key_b if self.pick_mode.get() == "manual" else None,
                    out_path=self.out_path.get(),
                    case_insensitive=self.case_insensitive.get(),
                    keep_duplicates=self.keep_duplicates.get(),
                    keep_blanks=self.keep_blanks.get(),
                )
                self.status.config(text=f"تمام شد | Matched={res['matched']} OnlyInA={res['only_a']} OnlyInB={res['only_b']}")
                messagebox.showinfo("تمام شد", f"خروجی Compare ساخته شد:\n{res['out']}")
                return

            if self.action.get() == "lookup":
                sel_idx = list(self.bcols_list.curselection())
                selected_cols = [self.bcols_list.get(i) for i in sel_idx] if sel_idx else None
                res = xlookup_join(
                    file_a=self.file_a.get(),
                    file_b=self.file_b.get(),
                    sheet_a=self.sheet_a.get() or None,
                    sheet_b=self.sheet_b.get() or None,
                    col_a=key_a if self.pick_mode.get() == "manual" else None,
                    col_b=key_b if self.pick_mode.get() == "manual" else None,
                    b_return_cols=selected_cols,
                    out_path=self.out_path.get(),
                    case_insensitive=self.case_insensitive.get(),
                    keep_blanks=self.keep_blanks.get(),
                )
                self.status.config(text=f"تمام شد | NotFound={res['not_found']} | DuplicatesInB={res['dup_rows_in_b']}")
                messagebox.showinfo("تمام شد", f"خروجی Lookup ساخته شد:\n{res['out']}")
                return

            # diff
            sel_idx = list(self.diffcols_list.curselection())
            selected_cols = [self.diffcols_list.get(i) for i in sel_idx] if sel_idx else None
            res = differences_report(
                file_a=self.file_a.get(),
                file_b=self.file_b.get(),
                sheet_a=self.sheet_a.get() or None,
                sheet_b=self.sheet_b.get() or None,
                col_a=key_a if self.pick_mode.get() == "manual" else None,
                col_b=key_b if self.pick_mode.get() == "manual" else None,
                compare_cols=selected_cols,
                out_path=self.out_path.get(),
                case_insensitive=self.case_insensitive.get(),
                keep_blanks=self.keep_blanks.get(),
            )
            self.status.config(text=f"تمام شد | Differences={res['differences']} Same={res['same']} NotFound={res['not_found']}")
            messagebox.showinfo("تمام شد", f"خروجی Differences ساخته شد:\n{res['out']}")

        except Exception as e:
            self.status.config(text="خطا")
            messagebox.showerror("خطا", str(e))


if __name__ == "__main__":
    App().mainloop()
