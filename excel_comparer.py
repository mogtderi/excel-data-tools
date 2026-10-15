import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class DataProcessor:
    @staticmethod
    def normalize_text(val):
        """نرمال‌سازی متون و اعداد برای مقایسه دقیق"""
        if pd.isna(val) or val == "":
            return ""
        val = str(val).strip()
        # رفع مشکل خوانده شدن اعداد صحیح به عنوان اعشاری در پانداس (مثل 505.0)
        if val.endswith('.0'):
            val = val[:-2]
            
        # یکسان‌سازی ی و ک
        val = val.replace('ي', 'ی').replace('ك', 'ک')
        # تبدیل اعداد فارسی به انگلیسی
        translation_table = str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')
        val = val.translate(translation_table)
        return val

class ExcelComparerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ابزار هوشمند مقایسه اکسل - نمایش قطعات گمشده با ❌ | Coded by Moghtaderi ")
        self.root.geometry("1000x700")
        self.root.configure(padx=20, pady=20)

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        
        self.df_result = None
        self.stats = {"total_checked": 0, "missing": 0}

        self.create_widgets()

    def create_widgets(self):
        # --- بخش انتخاب فایل‌ها ---
        frame_files = tk.LabelFrame(self.root, text="انتخاب فایل‌ها", padx=10, pady=10)
        frame_files.pack(fill="x", pady=(0, 10))

        tk.Label(frame_files, text=":فایل اصلی (لیست مرجع/شاهیندژ)").grid(row=0, column=2, sticky="e", pady=5)
        tk.Entry(frame_files, textvariable=self.file1_path, width=65, state="readonly").grid(row=0, column=1, padx=10)
        tk.Button(frame_files, text="انتخاب فایل", command=lambda: self.browse_file(self.file1_path)).grid(row=0, column=0)

        tk.Label(frame_files, text=":فایل دوم (داده‌ها/مقایسه)").grid(row=1, column=2, sticky="e", pady=5)
        tk.Entry(frame_files, textvariable=self.file2_path, width=65, state="readonly").grid(row=1, column=1, padx=10)
        tk.Button(frame_files, text="انتخاب فایل", command=lambda: self.browse_file(self.file2_path)).grid(row=1, column=0)

        # --- بخش تطبیق ستون‌ها ---
        frame_mapping = tk.LabelFrame(self.root, text="نام ستون‌های کلیدی (برای جستجو)", padx=10, pady=10)
        frame_mapping.pack(fill="x", pady=(0, 10))

        tk.Label(frame_mapping, text="فایل اصلی (مرجع)").grid(row=0, column=1, pady=5)
        tk.Label(frame_mapping, text="فایل دوم").grid(row=0, column=0, pady=5)

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

        # --- دکمه‌های کنترل ---
        frame_controls = tk.Frame(self.root)
        frame_controls.pack(fill="x", pady=5)
        tk.Button(frame_controls, text="شروع مقایسه دقیق", command=self.process_files, bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), width=15).pack(side="right", padx=5)
        tk.Button(frame_controls, text="خروجی اکسل با فرمت ❌", command=self.export_to_excel, bg="#f44336", fg="white", font=("Arial", 11, "bold"), width=25).pack(side="right", padx=5)

        # --- آمار ---
        self.lbl_stats = tk.Label(self.root, text="وضعیت: آماده برای مقایسه", font=("Arial", 11, "bold"), fg="#333")
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
                return pd.read_csv(filepath, dtype=str) # خواندن همه چیز به عنوان متن برای جلوگیری از خرابی داده
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

        # خواندن اسامی ستون‌ها از رابط گرافیکی
        village_col_f1 = self.col2_f1.get().strip()
        piece_base_col_f1 = self.col1_f1.get().strip().replace(" ", "")
        
        village_col_f2 = self.col2_f2.get().strip()
        piece_col_f2 = self.col1_f2.get().strip()

        # بررسی وجود ستون‌ها در فایل دوم
        if village_col_f2 not in df2.columns or piece_col_f2 not in df2.columns:
            messagebox.showerror("خطا", f"ستون‌های '{village_col_f2}' یا '{piece_col_f2}' در فایل دوم یافت نشد!")
            return

        # بررسی وجود ستون روستا در فایل اول
        if village_col_f1 not in df1.columns:
            messagebox.showerror("خطا", f"ستون '{village_col_f1}' در فایل مرجع یافت نشد!")
            return

        # پیدا کردن تمام ستون‌هایی که شامل کلمه "شماره قطعه" هستند (چون در فایل شما تکرار شده است)
        piece_columns_f1 = [col for col in df1.columns if piece_base_col_f1 in str(col).replace(" ", "")]
        
        if not piece_columns_f1:
            messagebox.showerror("خطا", f"هیچ ستونی با نام شبیه به '{self.col1_f1.get()}' در فایل مرجع یافت نشد!")
            return

        # مرحله ۱: ایجاد بانک اطلاعاتی سریع از فایل دوم
        # یک مجموعه (Set) از ترکیب (نام روستا, شماره قطعه) می‌سازیم تا جستجو در کسری از ثانیه انجام شود
        file2_records = set()
        for _, row in df2.iterrows():
            v = DataProcessor.normalize_text(row.get(village_col_f2, ""))
            p = DataProcessor.normalize_text(row.get(piece_col_f2, ""))
            if v and p:
                file2_records.add((v, p))

        # مرحله ۲: پیمایش فایل اول و جایگذاری علامت ❌
        self.df_result = df1.copy()
        self.stats = {"total_checked": 0, "missing": 0}

        for idx, row in self.df_result.iterrows():
            village = DataProcessor.normalize_text(row.get(village_col_f1, ""))
            if not village:
                continue
                
            # بررسی تک تک ستون‌های شماره قطعه در آن ردیف
            for p_col in piece_columns_f1:
                val = row[p_col]
                norm_val = DataProcessor.normalize_text(val)
                
                if norm_val: # اگر سلول خالی نبود
                    self.stats["total_checked"] += 1
                    
                    # اگر قطعه در فایل دوم نبود، علامت ❌ را به مقدار آن می‌چسبانیم
                    if (village, norm_val) not in file2_records:
                        # نمایش مقدار اصلی + ❌
                        original_clean = str(val).strip()
                        if original_clean.endswith('.0'):
                            original_clean = original_clean[:-2]
                        self.df_result.at[idx, p_col] = f"{original_clean} ❌"
                        self.stats["missing"] += 1

        # بروزرسانی رابط کاربری
        self.lbl_stats.config(
            text=f"📊 تعداد کل قطعات بررسی شده: {self.stats['total_checked']}  |  "
                 f"❌ قطعات مفقود شده (پیدا نشد): {self.stats['missing']}",
            fg="#d32f2f" if self.stats['missing'] > 0 else "#388e3c"
        )

        self.update_treeview(self.df_result)
        messagebox.showinfo("عملیات موفق", "مقایسه تمام شد! قطعاتی که در فایل دوم نبودند با علامت ❌ مشخص شدند.")

    def update_treeview(self, df):
        self.tree.delete(*self.tree.get_children())
        columns = list(df.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, anchor="center")

        for _, row in df.iterrows():
            values = [row[col] for col in columns]
            self.tree.insert("", "end", values=values)

    def export_to_excel(self):
        if self.df_result is None or self.df_result.empty:
            messagebox.showwarning("اخطار", "ابتدا باید فایل‌ها را مقایسه کنید.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Final_Report.xlsx",
            title="ذخیره فایل اکسل نهایی",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            try:
                # تابع برای رنگ‌آمیزی سلول‌هایی که دارای علامت ❌ هستند
                def style_missing(val):
                    if isinstance(val, str) and '❌' in val:
                        return 'background-color: #ffcccc; color: #cc0000; font-weight: bold;'
                    return ''
                
                # اعمال استایل به دیتافریم
                try:
                    styled_df = self.df_result.style.map(style_missing) # برای پانداس نسخه‌های جدید
                except AttributeError:
                    styled_df = self.df_result.style.applymap(style_missing) # برای پانداس نسخه‌های قدیمی‌تر
                
                styled_df.to_excel(save_path, index=False, engine='openpyxl')
                messagebox.showinfo("موفقیت", f"فایل با موفقیت ذخیره شد.\nقطعات گمشده در فایل اکسل با رنگ قرمز هایلایت شده‌اند.")
            except Exception as e:
                messagebox.showerror("خطا", f"خطا در ذخیره فایل:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparerApp(root)
    root.mainloop()