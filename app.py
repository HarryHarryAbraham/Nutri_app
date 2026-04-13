import customtkinter as ctk
from tkinter import messagebox, filedialog
from datetime import date, datetime
import os
import data_manager as dm


# ألوان
GRAY_BG = "#E0E0E0"
WHITE = "#FFFFFF"
BLUE_BTN = "#1E88E5"
GREEN_BTN = "#43A047"
ORANGE_BTN = "#FB8C00"
RED_BTN = "#E53935"

CATEGORIES = [
    "Children 6-59 months",
    "Children under 6 months",
    "Lactating 0-6 months",
    "Lactating 6-12 months",
    "Lactating 12-24 months",
    "Pregnant",
]
GENDERS = ["Male", "Female"]
IDP_OPTIONS = ["HOST", "IDP"]
ZSCORE_OPTIONS = ["(-3) to (-2)", "over (-2)", "under (-3)"]
MUAC_OPTIONS = [">135", "125 to 135", "115 to 125", "<115", "<230 for women", ">230 for women"]
DIAGNOSIS_OPTIONS = ["Normal", "Risk for acute malnutrition", "MAM", "SAM"]
ACTION_OPTIONS = ["Nothing", "Referred for Consultation", "Referred for Supplementation", "Referred for Treatment"]


class NutritionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("نظام تتبع التغذية - Nutrition Tracker")
        self.root.geometry("950x750")
        self.root.resizable(True, True)

        # متغيرات الحالة
        self.wb = None
        self.filepath = ""
        self.is_updating = False
        self.orig_mother_age = ""
        self.orig_child_age = ""

        # بناء الواجهة
        self._build_ui()

    # ─────────────────────────────────────────────
    #  بناء الواجهة الكاملة
    # ─────────────────────────────────────────────
    def _build_ui(self):
        # شريط العنوان
        title_bar = ctk.CTkFrame(self.root, fg_color="#1565C0", corner_radius=0)
        title_bar.pack(fill="x")
        ctk.CTkLabel(
            title_bar,
            text="🏥  نظام تتبع التغذية  |  Nutrition Tracker",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white",
        ).pack(pady=10)

        # زر فتح الملف
        file_frame = ctk.CTkFrame(self.root, fg_color="#E3F2FD", corner_radius=0)
        file_frame.pack(fill="x", padx=0, pady=0)

        ctk.CTkButton(
            file_frame,
            text="📂  فتح ملف Nutrition_Data.xlsx",
            command=self._open_file,
            fg_color=BLUE_BTN,
            width=260,
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left", padx=10, pady=8)

        self.lbl_file = ctk.CTkLabel(
            file_frame,
            text="لم يتم فتح أي ملف بعد",
            text_color="#555",
            font=ctk.CTkFont(size=12),
        )
        self.lbl_file.pack(side="left", padx=10)

        # التبويبات
        self.tabview = ctk.CTkTabview(self.root, anchor="nw")
        self.tabview.pack(fill="both", expand=True, padx=10, pady=5)

        self.tabview.add("📋  تسجيل الزيارة")
        self.tabview.add("🔍  بحث محلي")

        self._build_visit_tab(self.tabview.tab("📋  تسجيل الزيارة"))
        self._build_local_tab(self.tabview.tab("🔍  بحث محلي"))

    # ─────────────────────────────────────────────
    #  تبويب تسجيل الزيارة
    # ─────────────────────────────────────────────
    def _build_visit_tab(self, parent):
        # إطار التمرير
        scroll = ctk.CTkScrollableFrame(parent)
        scroll.pack(fill="both", expand=True, padx=5, pady=5)

        # ─ صف 1: المركز والتاريخ
        row1 = ctk.CTkFrame(scroll, fg_color="transparent")
        row1.pack(fill="x", pady=4)
        self._label(row1, "المركز الصحي:").pack(side="left", padx=(5, 2))
        self.txt_center = self._entry(row1, width=130)
        self.txt_center.insert(0, "Sosa")
        self.txt_center.pack(side="left", padx=5)

        self._label(row1, "التاريخ:").pack(side="left", padx=(15, 2))
        self.txt_date = self._entry(row1, width=120)
        self.txt_date.insert(0, date.today().strftime("%d/%m/%Y"))
        self.txt_date.pack(side="left", padx=5)

        # ─ صف 2: كود المريض
        row2 = ctk.CTkFrame(scroll, fg_color="transparent")
        row2.pack(fill="x", pady=4)
        self._label(row2, "كود المريض:").pack(side="left", padx=(5, 2))
        self.txt_code = self._entry(row2, width=160)
        self.txt_code.pack(side="left", padx=5)
        self.txt_code.bind("<Return>", lambda e: self._search_or_new())
        self.txt_code.bind("<KeyPress>", self._numbers_only_code)

        ctk.CTkButton(
            row2, text="🔎 بحث", command=self._search_or_new,
            fg_color=BLUE_BTN, width=90,
        ).pack(side="left", padx=5)

        # ─ صف 3: الفئة والجنس
        row3 = ctk.CTkFrame(scroll, fg_color="transparent")
        row3.pack(fill="x", pady=4)
        self._label(row3, "الفئة:").pack(side="left", padx=(5, 2))
        self.cmb_category = self._combo(row3, CATEGORIES, width=220)
        self.cmb_category.pack(side="left", padx=5)
        self.cmb_category.bind("<<ComboboxSelected>>", self._on_category_change)

        self._label(row3, "الجنس:").pack(side="left", padx=(15, 2))
        self.cmb_gender = self._combo(row3, GENDERS, width=110)
        self.cmb_gender.pack(side="left", padx=5)

        # ─ صف 4: عمر الأم / عمر الطفل
        row4 = ctk.CTkFrame(scroll, fg_color="transparent")
        row4.pack(fill="x", pady=4)
        self._label(row4, "عمر الأم (سنة):").pack(side="left", padx=(5, 2))
        self.txt_mother_age = self._entry(row4, width=100)
        self.txt_mother_age.pack(side="left", padx=5)
        self.txt_mother_age.bind("<KeyPress>", self._numbers_only_mother)

        self._label(row4, "عمر الطفل (شهر):").pack(side="left", padx=(15, 2))
        self.txt_child_age = self._entry(row4, width=100)
        self.txt_child_age.pack(side="left", padx=5)
        self.txt_child_age.bind("<KeyPress>", self._numbers_only_child)

        # ─ صف 5: IDP
        row5 = ctk.CTkFrame(scroll, fg_color="transparent")
        row5.pack(fill="x", pady=4)
        self._label(row5, "IDP / Host:").pack(side="left", padx=(5, 2))
        self.cmb_idp = self._combo(row5, IDP_OPTIONS, width=120)
        self.cmb_idp.set("HOST")
        self.cmb_idp.pack(side="left", padx=5)

        # ─ صف 6: MUAC و Z-Score
        row6 = ctk.CTkFrame(scroll, fg_color="transparent")
        row6.pack(fill="x", pady=4)
        self._label(row6, "MUAC (mm):").pack(side="left", padx=(5, 2))
        self.cmb_muac = self._combo(row6, MUAC_OPTIONS, width=160)
        self.cmb_muac.pack(side="left", padx=5)
        self.cmb_muac.bind("<<ComboboxSelected>>", self._on_muac_change)

        self._label(row6, "Z-Score:").pack(side="left", padx=(15, 2))
        self.cmb_zscore = self._combo(row6, ZSCORE_OPTIONS, width=150)
        self.cmb_zscore.pack(side="left", padx=5)

        # ─ صف 7: التشخيص والإجراء
        row7 = ctk.CTkFrame(scroll, fg_color="transparent")
        row7.pack(fill="x", pady=4)
        self._label(row7, "التشخيص:").pack(side="left", padx=(5, 2))
        self.cmb_diagnosis = self._combo(row7, DIAGNOSIS_OPTIONS, width=220)
        self.cmb_diagnosis.pack(side="left", padx=5)

        self._label(row7, "الإجراء:").pack(side="left", padx=(15, 2))
        self.cmb_action = self._combo(row7, ACTION_OPTIONS, width=230)
        self.cmb_action.pack(side="left", padx=5)

        # ─ فاصل
        ctk.CTkFrame(scroll, height=2, fg_color="#BDBDBD").pack(fill="x", padx=5, pady=10)

        # ─ أزرار العمليات
        btn_row = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_row.pack(fill="x", pady=5)

        ctk.CTkButton(
            btn_row, text="💾  حفظ الزيارة", command=self._save_visit,
            fg_color=GREEN_BTN, font=ctk.CTkFont(size=13, weight="bold"), width=160,
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            btn_row, text="✏️  تعديل البيانات الثابتة", command=self._edit_fixed,
            fg_color=ORANGE_BTN, width=190,
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            btn_row, text="🗑️  تفريغ الحقول", command=self._clear_fields,
            fg_color="#757575", width=140,
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            btn_row, text="📤  تصدير Excel", command=self._export,
            fg_color="#6A1B9A", width=140,
        ).pack(side="left", padx=8)

        # ─ شريط الحالة
        self.lbl_status = ctk.CTkLabel(
            scroll,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="#1565C0",
        )
        self.lbl_status.pack(pady=5)

    # ─────────────────────────────────────────────
    #  تبويب البحث المحلي
    # ─────────────────────────────────────────────
    def _build_local_tab(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # صف البحث
        search_row = ctk.CTkFrame(frame, fg_color="transparent")
        search_row.pack(fill="x", pady=8)
        self._label(search_row, "رقم المريض:").pack(side="left", padx=(5, 2))
        self.txt_patient_num = self._entry(search_row, width=160)
        self.txt_patient_num.pack(side="left", padx=5)
        self.txt_patient_num.bind("<Return>", lambda e: self._search_local())

        ctk.CTkButton(
            search_row, text="🔎 بحث", command=self._search_local,
            fg_color=BLUE_BTN, width=90,
        ).pack(side="left", padx=5)

        # نتائج البحث
        result_frame = ctk.CTkFrame(frame, fg_color="#F5F5F5", corner_radius=10)
        result_frame.pack(fill="x", pady=10, padx=5)

        fields = [
            ("اسم المريض الكامل:", "txt_patient_name"),
            ("تاريخ الميلاد:", "txt_birthdate"),
            ("العمر:", "txt_age_x"),
        ]
        for label_text, attr_name in fields:
            r = ctk.CTkFrame(result_frame, fg_color="transparent")
            r.pack(fill="x", padx=10, pady=4)
            self._label(r, label_text, width=150).pack(side="left")
            entry = self._entry(r, width=380)
            entry.configure(state="readonly")
            entry.pack(side="left", padx=5)
            setattr(self, attr_name, entry)

    # ─────────────────────────────────────────────
    #  منطق البحث الرئيسي
    # ─────────────────────────────────────────────
    def _search_or_new(self):
        if not self._check_file():
            return
        code = self.txt_code.get().strip()
        if not code:
            return

        self.is_updating = True
        row = dm.search_patient(self.wb, code)

        if row:
            self._clear_keep_code()
            # حساب الفترة الزمنية
            last_date = row[1]
            if isinstance(last_date, datetime):
                last_date = last_date.date()
            elif isinstance(last_date, str):
                try:
                    last_date = datetime.strptime(last_date, "%d/%m/%Y").date()
                except Exception:
                    last_date = date.today()

            today = date.today()
            months_diff = (today.year - last_date.year) * 12 + (today.month - last_date.month)
            years_diff = today.year - last_date.year

            # عمر الأم
            if row[5] is not None and str(row[5]).strip() != "":
                try:
                    new_mother = int(float(str(row[5]))) + years_diff
                    self._set_entry(self.txt_mother_age, str(new_mother))
                    self.orig_mother_age = str(new_mother)
                except Exception:
                    pass

            # عمر الطفل
            new_child_age = None
            if row[6] is not None and str(row[6]).strip() != "":
                try:
                    new_child_age = int(float(str(row[6]))) + months_diff
                    self._set_entry(self.txt_child_age, str(new_child_age))
                    self.orig_child_age = str(new_child_age)
                except Exception:
                    pass

            # تحديث الفئة
            old_cat = str(row[3]) if row[3] else ""
            if "Children" in old_cat and new_child_age is not None:
                if 0 <= new_child_age <= 5:
                    self.cmb_category.set("Children under 6 months")
                elif 6 <= new_child_age <= 59:
                    self.cmb_category.set("Children 6-59 months")
                self.cmb_category.configure(state="disabled")
                self.cmb_muac.focus_set()

            elif "Lactating" in old_cat and new_child_age is not None:
                if 0 <= new_child_age <= 6:
                    self.cmb_category.set("Lactating 0-6 months")
                elif 7 <= new_child_age <= 12:
                    self.cmb_category.set("Lactating 6-12 months")
                elif 13 <= new_child_age <= 24:
                    self.cmb_category.set("Lactating 12-24 months")
                self.cmb_category.configure(state="normal")
                self.cmb_category.focus_set()

            elif old_cat == "Pregnant":
                self.cmb_category.set("Pregnant")
                self.cmb_category.configure(state="normal")
                self.cmb_category.focus_set()

            # الجنس والإقامة
            if row[4]:
                self.cmb_gender.set(str(row[4]))
            if row[7]:
                self.cmb_idp.set(str(row[7]))

            self._lock_fixed(True)
            self._set_status(f"✅ تم استدعاء بيانات المريض رقم {code}", "green")

        else:
            self._clear_keep_code()
            self._lock_fixed(False)
            messagebox.showinfo(
                "مريض جديد",
                f"رقم المريض ({code}) غير موجود في قاعدة البيانات.\nسيتم تسجيله كمريض جديد.",
            )
            self.cmb_category.focus_set()

        self.is_updating = False

    # ─────────────────────────────────────────────
    #  حدث تغيير الفئة
    # ─────────────────────────────────────────────
    def _on_category_change(self, event=None):
        if self.is_updating:
            return
        cat = self.cmb_category.get()

        if "Children" in cat:
            self._set_entry(self.txt_mother_age, "")
            self.txt_mother_age.configure(state="disabled", fg_color=GRAY_BG)
            if self.orig_child_age:
                self._set_entry(self.txt_child_age, self.orig_child_age)
            self.txt_child_age.configure(state="normal", fg_color=WHITE)
            self.cmb_gender.set("")
            if cat == "Children under 6 months":
                self.cmb_muac.set("")
                self.cmb_zscore.set("over (-2)")
                self.cmb_diagnosis.set("Normal")
                self.cmb_action.set("Nothing")

        elif cat == "Pregnant":
            self._set_entry(self.txt_child_age, "")
            self.txt_child_age.configure(state="disabled", fg_color=GRAY_BG)
            if self.orig_mother_age:
                self._set_entry(self.txt_mother_age, self.orig_mother_age)
            self.txt_mother_age.configure(state="normal", fg_color=WHITE)
            self.cmb_gender.set("Female")

        elif "Lactating" in cat:
            if self.orig_mother_age:
                self._set_entry(self.txt_mother_age, self.orig_mother_age)
            if self.orig_child_age:
                self._set_entry(self.txt_child_age, self.orig_child_age)
            self.txt_child_age.configure(state="normal", fg_color=WHITE)
            self.txt_mother_age.configure(state="normal", fg_color=WHITE)
            self.cmb_gender.set("Female")

    # ─────────────────────────────────────────────
    #  حدث تغيير MUAC
    # ─────────────────────────────────────────────
    def _on_muac_change(self, event=None):
        if self.is_updating:
            return
        muac = self.cmb_muac.get()
        mapping = {
            ">135":           ("over (-2)",    "Normal",                    "Nothing"),
            "125 to 135":     ("over (-2)",    "Risk for acute malnutrition","Referred for Consultation"),
            "115 to 125":     ("(-3) to (-2)", "MAM",                       "Referred for Supplementation"),
            "<115":           ("under (-3)",   "SAM",                       "Referred for Treatment"),
            "<230 for women": ("",             "MAM",                       "Referred for Supplementation"),
            ">230 for women": ("",             "Normal",                    "Nothing"),
        }
        if muac in mapping:
            z, diag, action = mapping[muac]
            self.cmb_zscore.set(z)
            self.cmb_diagnosis.set(diag)
            self.cmb_action.set(action)

    # ─────────────────────────────────────────────
    #  حفظ الزيارة
    # ─────────────────────────────────────────────
    def _save_visit(self):
        if not self._check_file():
            return
        code = self.txt_code.get().strip()
        category = self.cmb_category.get()
        if not code or not category:
            messagebox.showerror("نقص بيانات", "يرجى تعبئة الكود والفئة على الأقل.")
            return

        # تحويل التاريخ
        try:
            d = datetime.strptime(self.txt_date.get().strip(), "%d/%m/%Y").date()
        except Exception:
            d = date.today()

        mother_val = self.txt_mother_age.get().strip()
        child_val = self.txt_child_age.get().strip()

        data = {
            "center":     self.txt_center.get().strip(),
            "date":       d,
            "code":       code,
            "category":   category,
            "gender":     self.cmb_gender.get(),
            "mother_age": int(mother_val) if mother_val.isdigit() else None,
            "child_age":  int(child_val) if child_val.isdigit() else None,
            "idp":        self.cmb_idp.get(),
            "zscore":     self.cmb_zscore.get(),
            "muac":       self.cmb_muac.get(),
            "diagnosis":  self.cmb_diagnosis.get(),
            "action":     self.cmb_action.get(),
        }

        try:
            dm.save_visit(self.wb, data)
            dm.save_file(self.wb, self.filepath)
            messagebox.showinfo("نجاح", "✅ تم حفظ الزيارة بنجاح!")
            self._set_entry(self.txt_code, "")
            self._clear_keep_code()
            self.txt_code.focus_set()
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل الحفظ:\n{e}")

    # ─────────────────────────────────────────────
    #  تصدير Excel
    # ─────────────────────────────────────────────
    def _export(self):
        if not self._check_file():
            return
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        filename = f"Nutrition_Exported_{datetime.now().strftime('%Y_%m_%d_%H%M%S')}.xlsx"
        dest = os.path.join(desktop, filename)
        try:
            dm.export_nutrition_data(self.wb, dest)
            messagebox.showinfo("تصدير", f"✅ تم التصدير بنجاح إلى:\n{dest}")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")

    # ─────────────────────────────────────────────
    #  البحث المحلي
    # ─────────────────────────────────────────────
    def _search_local(self):
        if not self._check_file():
            return
        num = self.txt_patient_num.get().strip()
        if not num:
            return

        result = dm.search_local(self.wb, num)
        if result:
            full_name = " ".join(filter(None, [
                str(result["first_name"]),
                str(result["father_name"]),
                str(result["last_name"]),
            ]))
            self._set_readonly(self.txt_patient_name, full_name)

            bd = result["birthdate"]
            if isinstance(bd, datetime):
                bd_str = bd.strftime("%d/%m/%Y")
            elif isinstance(bd, str):
                bd_str = bd
            else:
                bd_str = ""
            self._set_readonly(self.txt_birthdate, bd_str)

            age_str = dm.calculate_age(result["birthdate"])
            self._set_readonly(self.txt_age_x, age_str)
        else:
            messagebox.showwarning("غير موجود", f"الرقم ({num}) غير موجود محلياً!")
            self._set_readonly(self.txt_patient_name, "")
            self._set_readonly(self.txt_birthdate, "")
            self._set_readonly(self.txt_age_x, "")

    # ─────────────────────────────────────────────
    #  فتح الملف
    # ─────────────────────────────────────────────
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="اختر ملف Nutrition_Data.xlsx",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            self.wb = dm.load_file(path)
            self.filepath = path
            self.lbl_file.configure(text=f"✅ {os.path.basename(path)}", text_color="green")
            self._set_status("تم فتح الملف بنجاح. ضع كود المريض وانقر بحث.", "#1565C0")
            self.txt_code.focus_set()
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل فتح الملف:\n{e}")

    # ─────────────────────────────────────────────
    #  دوال مساعدة
    # ─────────────────────────────────────────────
    def _check_file(self):
        if not self.wb:
            messagebox.showwarning("تنبيه", "يرجى فتح ملف Excel أولاً.")
            return False
        return True

    def _clear_keep_code(self):
        self.is_updating = True
        self.cmb_category.set("")
        self.cmb_category.configure(state="normal")
        self.cmb_gender.set("")
        self._set_entry(self.txt_mother_age, "")
        self._set_entry(self.txt_child_age, "")
        self.cmb_idp.set("HOST")
        self.cmb_zscore.set("")
        self.cmb_muac.set("")
        self.cmb_diagnosis.set("")
        self.cmb_action.set("")
        self.orig_mother_age = ""
        self.orig_child_age = ""
        # إعادة تفعيل الحقول
        for w in [self.txt_mother_age, self.txt_child_age]:
            w.configure(state="normal", fg_color=WHITE)
        for w in [self.cmb_gender, self.cmb_idp]:
            w.configure(state="normal", fg_color=WHITE)
        self.is_updating = False

    def _clear_fields(self):
        self.is_updating = True
        self._set_entry(self.txt_code, "")
        self._clear_keep_code()
        self.is_updating = False
        self.txt_code.focus_set()

    def _lock_fixed(self, lock: bool):
        state = "disabled" if lock else "normal"
        color = GRAY_BG if lock else WHITE
        self.cmb_gender.configure(state=state, fg_color=color)
        self.cmb_idp.configure(state=state, fg_color=color)

    def _edit_fixed(self):
        self._lock_fixed(False)
        messagebox.showinfo("تعديل", "يمكنك الآن تعديل البيانات الثابتة.")

    def _set_status(self, msg, color="#1565C0"):
        self.lbl_status.configure(text=msg, text_color=color)

    def _set_entry(self, widget, value):
        widget.configure(state="normal")
        widget.delete(0, "end")
        widget.insert(0, value)

    def _set_readonly(self, widget, value):
        widget.configure(state="normal")
        widget.delete(0, "end")
        widget.insert(0, value)
        widget.configure(state="readonly")

    # ─── قيود الإدخال الرقمي ───
    def _numbers_only_code(self, event):
        self._allow_numbers(event, self.txt_code)

    def _numbers_only_mother(self, event):
        self._allow_numbers(event, self.txt_mother_age)

    def _numbers_only_child(self, event):
        self._allow_numbers(event, self.txt_child_age)

    def _allow_numbers(self, event, widget):
        if event.keysym in ("BackSpace", "Delete", "Left", "Right", "Tab", "Return"):
            return
        if not event.char.isdigit():
            return "break"

    # ─── widgets مساعدة ───
    def _label(self, parent, text, width=None):
        kwargs = {"text": text, "font": ctk.CTkFont(size=13), "anchor": "w"}
        if width:
            kwargs["width"] = width
        return ctk.CTkLabel(parent, **kwargs)

    def _entry(self, parent, width=180):
        return ctk.CTkEntry(parent, width=width, font=ctk.CTkFont(size=13))

    def _combo(self, parent, values, width=180):
        import tkinter.ttk as ttk
        combo = ttk.Combobox(parent, values=values, width=width // 8, font=("Arial", 12))
        combo.configure(state="normal")
        return combo
