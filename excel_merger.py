# ==========================================
# Coded by Moghtaderi
# Smart Excel Merger & Missing Records Finder
# ==========================================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class DataProcessor:
    @staticmethod
    def normalize_text(val):
        """نرمال‌سازی متون و اعداد برای جلوگیری از عدم تطابق به دلیل فاصله‌ها یا حروف عربی"""
        if pd.isna(val) or val == "":
            return ""
        val = str(val).strip()
        if val.endswith('.0'):
            val = val[:-2]
        val = val.replace('ي', 'ی').replace('ك', 'ک')
        translation_table = str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')
        val = val.translate(translation_table)
        return val

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        # اضافه شدن نام برنامه‌نویس به عنوان پنجره
        self.root.title("ابزار ادغام هوشمند لیست‌های اکسل | Coded by Moghtaderi")
        self.root.geometry("1050x750")
        self.root.configure(padx=20, pady=20)

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        
        self.df_merged = None
        self.stats = {"total": 0, "both": 0, "only_f1": 0, "only_f2": 0}

        self.create_widgets()

    def create_widgets(self):
        # --- بخش انتخاب فایل‌ها ---
        frame_files = tk.LabelFrame(self.root, text="انتخاب فایل‌ها", padx=10, pady=10)
        frame_files.pack(fill="x", pady=(0, 10))

        tk.Label(frame_files, text=":فایل اصلی (لیست مرجع/شاهیندژ)").grid(row=0, column=2, sticky="e", pady=5)
        tk.Entry(frame_files, textvariable=self.file1_path, width=65, state="readonly").grid(row=0, column=1, padx=10)
        tk.Button(frame_files, text="انتخاب فایل", command=lambda: self.browse_file(self.file1_path)).grid(row=0, column=0)

        tk.Label(frame_files, text=":فایل دوم (داده‌های تکمیلی/Report)").grid(row=1, column=2, sticky="e", pady=5)
        tk.Entry(frame_files, textvariable=self.file2_path, width=65, state="readonly").grid(row=1, column=1, padx=10)
        tk.Button(frame_files, text="انتخاب فایل", command=lambda: self.browse_file(self.file2_path)).grid(row=1, column=0)

        # --- بخش تطبیق ستون‌ها ---
        frame_mapping = tk.LabelFrame(self.root, text="تنظیمات ستون‌ها", padx=10, pady=10)
        frame_mapping.pack(fill="x", pady=(0, 10))

        tk.Label(frame_mapping, text="فایل اصلی (مرجع)").grid(row=0, column=1, pady=5)
        tk.Label(frame_mapping, text="فایل دوم (تکمیلی)").grid(row=0, column=0, pady=5)

        tk.Label(frame_mapping, text=":ستون شماره قطعه").grid(row=1, column=2, sticky="e")
        self.col1_f1 = tk.Entry(frame_mapping, justify="center")
        self.col1_f1.insert(0, "شماره قطعه")
        self.col1_f1.grid(row=1, column=1, padx=5, pady=2)
        
        self.col1_f2 = tk.Entry(frame_mapping, justify="center")
        self.col1_f2.insert(0, "شماره قطعه")
        self.col1_f2.grid(row=1, column=0, padx=5, pady=2)

        tk.Label(frame_mapping, text=":ستون نام روستا/آبادی").grid(row=2, column=2, sticky="e")
        self.col2_f1 = tk.Entry(frame_mapping, justify="center")
        self.col2_f1.insert(0, "نام روستا")
        self.col2_f1.grid(row=2, column=1, padx=5, pady=2)
        
        self.col2_f2 = tk.Entry(frame_mapping, justify="center")
        self.col2_f2.insert(0, "نام آبادی")
        self.col2_f2.grid(row=2, column=0, padx=5, pady=2)

        tk.Label(frame_mapping, text=":ستون مساحت (در فایل دوم)").grid(row=3, column=2, sticky="e")
        self.col_area = tk.Entry(frame_mapping, justify="center")
        self.col_area.insert(0, "مساحت")
        self.col_area.grid(row=3, column=0, padx=5, pady=2)

        tk.Label(frame_mapping, text=":ستون متقاضیان (در فایل دوم)").grid(row=4, column=2, sticky="e")
        self.col_app = tk.Entry(frame_mapping, justify="center")
        self.col_app.insert(0, "متقاضیان")
        self.col_app.grid(row=4, column=0, padx=5, pady=2)

        # --- دکمه‌های کنترل ---
        frame_controls = tk.Frame(self.root)
        frame_controls.pack(fill="x", pady=5)
        tk.Button(frame_controls, text="ادغام لیست‌ها (بدون تکرار)", command=self.process_files, bg="#2196F3", fg="white", font=("Arial", 11, "bold"), width=22).pack(side="right", padx=5)
        tk.Button(frame_controls, text="خروجی فایل نهایی", command=self.export_to_excel, bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), width=20).pack(side="right", padx=5)

        # --- آمار ---
        self.lbl_stats = tk.Label(self.root, text="وضعیت: آماده برای ادغام", font=("Arial", 10, "bold"), fg="#333")
        self.lbl_stats.pack(fill="x", pady=10)

        # --- جدول نمایش نتایج ---
        frame_tree = tk.Frame(self.root)
        frame_tree.pack(fill="both", expand=True)

        scroll_y = tk.Scrollbar(frame_tree, orient="vertical")
        scroll_x = tk.Scrollbar(frame_tree, orient="horizontal")
        
        self.tree = ttk.Treeview(frame_tree, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        # --- فوتر برنامه‌نویس ---
        lbl_footer = tk.Label(self.root, text="Coded by Moghtaderi", font=("Arial", 10, "italic"), fg="#7f8c8d")
        lbl_footer.pack(side="bottom", pady=(10, 0))

    def browse_file(self, string_var):
        filename = filedialog.askopenfilename(
            title="انتخاب فایل",
            filetypes=[("Excel & CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if filename:
            string_var.set(filename)

    def read_file(self, filepath):
        try:
            if filepath.endswith('.csv'):
                return pd.read_csv(filepath, dtype=str)
            else:
                return pd.read_excel(filepath, dtype=str)
        except Exception as e:
            messagebox.showerror("خطا", f"امکان خواندن فایل وجود ندارد:\n{str(e)}")
            return None

    def process_files(self):
        f1_path = self.file1_path.get()
        f2_path = self.file2_path.get()

        if not f1_path or not f2_path:
            messagebox.showwarning("خطا", "لطفاً هر دو فایل را انتخاب کنید.")
            return

        df1 = self.read_file(f1_path)
        df2 = self.read_file(f2_path)

        if df1 is None or df2 is None: return

        v_col_1, p_col_1 = self.col2_f1.get().strip(), self.col1_f1.get().strip()
        v_col_2, p_col_2 = self.col2_f2.get().strip(), self.col1_f2.get().strip()
        area_col, app_col = self.col_area.get().strip(), self.col_app.get().strip()

        if v_col_1 not in df1.columns:
            messagebox.showerror("خطا", f"ستون '{v_col_1}' در فایل اول یافت نشد.")
            return
            
        piece_columns_f1 = [col for col in df1.columns if p_col_1.replace(" ", "") in str(col).replace(" ", "")]
        if not piece_columns_f1:
            messagebox.showerror("خطا", f"هیچ ستونی شبیه به '{p_col_1}' در فایل اول یافت نشد.")
            return

        # ---- مرحله 1: تبدیل فایل اول (افقی) به لیست عمودی استاندارد ----
        df1_melted = pd.melt(df1, id_vars=[v_col_1], value_vars=piece_columns_f1, value_name="قطعه_موقت")
        df1_melted = df1_melted.dropna(subset=["قطعه_موقت"])
        
        df1_melted['روستا_نرمال'] = df1_melted[v_col_1].apply(DataProcessor.normalize_text)
        df1_melted['قطعه_نرمال'] = df1_melted["قطعه_موقت"].apply(DataProcessor.normalize_text)
        
        df1_clean = df1_melted[df1_melted['قطعه_نرمال'] != ""]
        df1_clean = df1_clean.drop_duplicates(subset=['روستا_نرمال', 'قطعه_نرمال'])
        df1_clean['حضور_در_فایل1'] = True

        # ---- مرحله 2: پاکسازی فایل دوم ----
        df2_clean = df2.copy()
        df2_clean['روستا_نرمال'] = df2_clean[v_col_2].apply(DataProcessor.normalize_text)
        df2_clean['قطعه_نرمال'] = df2_clean[p_col_2].apply(DataProcessor.normalize_text)
        df2_clean = df2_clean[df2_clean['قطعه_نرمال'] != ""]
        
        df2_clean = df2_clean.drop_duplicates(subset=['روستا_نرمال', 'قطعه_نرمال'])
        df2_clean['حضور_در_فایل2'] = True

        # ---- مرحله 3: ادغام لیست‌ها (Outer Join) ----
        merged = pd.merge(df1_clean, df2_clean, on=['روستا_نرمال', 'قطعه_نرمال'], how='outer')

        def get_status(row):
            if row.get('حضور_در_فایل1') == True and row.get('حضور_در_فایل2') == True:
                return 'موجود در هر دو'
            elif row.get('حضور_در_فایل1') == True:
                return 'فقط لیست اصلی'
            else:
                return 'فقط لیست تکمیلی'

        merged['وضعیت'] = merged.apply(get_status, axis=1)

        # ---- مرحله 4: ساخت ساختار نهایی خروجی ----
        final_data = {
            'نام روستا': merged['روستا_نرمال'],
            'شماره قطعه': merged['قطعه_نرمال'],
            'مساحت': merged[area_col] if area_col in df2_clean.columns else "نامشخص",
            'متقاضیان': merged[app_col] if app_col in df2_clean.columns else "نامشخص",
            'وضعیت رکورد': merged['وضعیت']
        }
        
        self.df_merged = pd.DataFrame(final_data)
        self.df_merged.fillna("-", inplace=True)
        
        self.df_merged.sort_values(by=['نام روستا', 'شماره قطعه'], inplace=True)

        self.stats['total'] = len(self.df_merged)
        self.stats['both'] = len(self.df_merged[self.df_merged['وضعیت رکورد'] == 'موجود در هر دو'])
        self.stats['only_f1'] = len(self.df_merged[self.df_merged['وضعیت رکورد'] == 'فقط لیست اصلی'])
        self.stats['only_f2'] = len(self.df_merged[self.df_merged['وضعیت رکورد'] == 'فقط لیست تکمیلی'])

        self.lbl_stats.config(
            text=f"کل رکوردهای یکتا: {self.stats['total']} | "
                 f"موجود در هر دو: {self.stats['both']} | "
                 f"فقط لیست اصلی: {self.stats['only_f1']} | "
                 f"فقط لیست تکمیلی: {self.stats['only_f2']}",
            fg="#000000"
        )

        self.update_treeview(self.df_merged)
        messagebox.showinfo("عملیات موفق", "ادغام لیست‌ها با موفقیت انجام شد.\nرکوردهای تکراری حذف شدند و اطلاعات تکمیلی اضافه شد.")

    def update_treeview(self, df):
        self.tree.delete(*self.tree.get_children())
        columns = list(df.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")

        for _, row in df.iterrows():
            values = [row[col] for col in columns]
            self.tree.insert("", "end", values=values)

    def export_to_excel(self):
        if self.df_merged is None or self.df_merged.empty:
            messagebox.showwarning("اخطار", "لیستی برای ذخیره وجود ندارد. ابتدا ادغام را انجام دهید.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Merged_Final_List.xlsx",
            title="ذخیره فایل نهایی ادغام شده",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            try:
                self.df_merged.to_excel(save_path, index=False, engine='openpyxl')
                messagebox.showinfo("موفقیت", "فایل ادغام شده با موفقیت ذخیره شد.")
            except Exception as e:
                messagebox.showerror("خطا", f"خطا در ذخیره فایل:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()