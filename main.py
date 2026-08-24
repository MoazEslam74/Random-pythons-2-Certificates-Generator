import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
import pythoncom
from docxtpl import DocxTemplate

# ==========================================
# قاموس اللغات لدعم العربية والإنجليزية
# ==========================================
LANG_DICT = {
    'EN': {
        'window_title': 'Certificate Generator',
        'lang_btn': 'عربي',
        'paths_frame': 'Files & Paths',
        'template_lbl': 'Word Template Path:',
        'output_lbl': 'Output Folder Path:',
        'browse_btn': 'Browse',
        'tags_frame': 'Tags Configuration',
        'tags_lbl': 'New Tag Name:',
        'add_tag_btn': '+ Add Tag',
        'current_tags': 'Current Tags: ',
        'input_frame': 'Data Entry (Lock 🔒 constant tags like Date/Manager)',
        'add_cert_btn': 'Add to List (or press Enter)',
        'list_frame': 'Certificates List',
        'generate_btn': '🚀 Generate Certificates',
        'status_ready': 'Status: Ready',
        'error_no_template': 'Please select a Word template path.',
        'error_no_output': 'Please select an output folder.',
        'error_no_certs': 'Please add at least one certificate.',
        'error_empty_fields': 'Please fill all fields before adding.',
        'error_tag_exists': 'Tag already exists.',
        'msg_done': '🎉 Certificates generated successfully!',
        'msg_generating': 'Generating... Please wait.',
        'del_btn': 'X'
    },
    'AR': {
        'window_title': 'مولد الشهادات الذكي',
        'lang_btn': 'English',
        'paths_frame': 'الملفات والمسارات',
        'template_lbl': 'مسار قالب الوورد:',
        'output_lbl': 'مسار مجلد الحفظ:',
        'browse_btn': 'تصفح',
        'tags_frame': 'إعدادات الـ Tags',
        'tags_lbl': 'اسم Tag جديد:',
        'add_tag_btn': '+ إضافة',
        'current_tags': 'الـ Tags الحالية: ',
        'input_frame': 'إدخال البيانات (اضغط 🔒 لتثبيت البيانات التي لا تتغير مثل التاريخ والمدير)',
        'add_cert_btn': 'إضافة للقائمة (أو اضغط Enter)',
        'list_frame': 'قائمة الشهادات',
        'generate_btn': '🚀 إصدار الشهادات',
        'status_ready': 'الحالة: مستعد',
        'error_no_template': 'يرجى تحديد مسار قالب الوورد.',
        'error_no_output': 'يرجى تحديد مسار مجلد الحفظ.',
        'error_no_certs': 'يرجى إضافة شهادة واحدة على الأقل للقائمة.',
        'error_empty_fields': 'يرجى تعبئة جميع الحقول قبل الإضافة.',
        'error_tag_exists': 'هذا الـ Tag موجود بالفعل.',
        'msg_done': '🎉 تم إصدار جميع الشهادات بنجاح!',
        'msg_generating': 'جاري الإصدار... يرجى الانتظار.',
        'del_btn': 'X'
    }
}

class CertificateApp:
    def __init__(self, root):
        self.root = root
        self.lang = 'AR'
        self.tags = ['name']
        self.locked_tags = {'name': False} # قاموس لتتبع حالة القفل لكل Tag
        self.certificates = []
        
        self.entry_widgets = {}
        self.lock_buttons = {} # لتخزين أزرار القفل للتحكم بها
        
        self.root.geometry("750x750")
        
        self.setup_ui()
        self.update_ui_texts()
        self.refresh_input_frame()
        self.refresh_tags_display()

    def setup_ui(self):
        self.lang_btn = ttk.Button(self.root, command=self.toggle_lang)
        self.lang_btn.pack(anchor='ne', padx=10, pady=5)

        self.paths_frame = ttk.LabelFrame(self.root)
        self.paths_frame.pack(fill='x', padx=10, pady=5)
        
        self.template_lbl = ttk.Label(self.paths_frame)
        self.template_lbl.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.template_path_var = tk.StringVar()
        ttk.Entry(self.paths_frame, textvariable=self.template_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        self.template_browse_btn = ttk.Button(self.paths_frame, command=self.browse_template)
        self.template_browse_btn.grid(row=0, column=2, padx=5, pady=5)

        self.output_lbl = ttk.Label(self.paths_frame)
        self.output_lbl.grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.output_path_var = tk.StringVar()
        ttk.Entry(self.paths_frame, textvariable=self.output_path_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        self.output_browse_btn = ttk.Button(self.paths_frame, command=self.browse_output)
        self.output_browse_btn.grid(row=1, column=2, padx=5, pady=5)

        self.tags_frame = ttk.LabelFrame(self.root)
        self.tags_frame.pack(fill='x', padx=10, pady=5)
        
        self.tag_input_lbl = ttk.Label(self.tags_frame)
        self.tag_input_lbl.grid(row=0, column=0, padx=5, pady=5)
        self.new_tag_var = tk.StringVar()
        tag_entry = ttk.Entry(self.tags_frame, textvariable=self.new_tag_var, width=20)
        tag_entry.grid(row=0, column=1, padx=5, pady=5)
        tag_entry.bind('<Return>', lambda e: self.add_tag())
        
        self.add_tag_btn = ttk.Button(self.tags_frame, command=self.add_tag)
        self.add_tag_btn.grid(row=0, column=2, padx=5, pady=5)
        
        self.current_tags_lbl = ttk.Label(self.tags_frame, foreground="blue")
        self.current_tags_lbl.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky='w')

        self.input_frame = ttk.LabelFrame(self.root)
        self.input_frame.pack(fill='x', padx=10, pady=5)
        self.dynamic_entries_frame = ttk.Frame(self.input_frame)
        self.dynamic_entries_frame.pack(fill='x', padx=5, pady=5)
        
        self.add_cert_btn = ttk.Button(self.input_frame, command=self.add_certificate)
        self.add_cert_btn.pack(pady=5)

        self.list_frame = ttk.LabelFrame(self.root)
        self.list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.canvas = tk.Canvas(self.list_frame)
        self.scrollbar = ttk.Scrollbar(self.list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_list = ttk.Frame(self.canvas)
        
        self.scrollable_list.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_list, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.bottom_frame = ttk.Frame(self.root)
        self.bottom_frame.pack(fill='x', padx=10, pady=10)
        
        self.generate_btn = ttk.Button(self.bottom_frame, command=self.start_generation)
        self.generate_btn.pack(side='left', padx=5)
        
        self.status_lbl = ttk.Label(self.bottom_frame, font=('Arial', 10, 'bold'))
        self.status_lbl.pack(side='left', padx=10)

    def update_ui_texts(self):
        t = LANG_DICT[self.lang]
        self.root.title(t['window_title'])
        self.lang_btn.config(text=t['lang_btn'])
        
        self.paths_frame.config(text=t['paths_frame'])
        self.template_lbl.config(text=t['template_lbl'])
        self.output_lbl.config(text=t['output_lbl'])
        self.template_browse_btn.config(text=t['browse_btn'])
        self.output_browse_btn.config(text=t['browse_btn'])
        
        self.tags_frame.config(text=t['tags_frame'])
        self.tag_input_lbl.config(text=t['tags_lbl'])
        self.add_tag_btn.config(text=t['add_tag_btn'])
        
        self.input_frame.config(text=t['input_frame'])
        self.add_cert_btn.config(text=t['add_cert_btn'])
        
        self.list_frame.config(text=t['list_frame'])
        self.generate_btn.config(text=t['generate_btn'])
        
        self.status_lbl.config(text=t['status_ready'])
        self.refresh_tags_display()

    def toggle_lang(self):
        self.lang = 'EN' if self.lang == 'AR' else 'AR'
        self.update_ui_texts()
        self.render_all_certificates()

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self.template_path_var.set(path)

    def browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_path_var.set(path)

    def add_tag(self):
        new_tag = self.new_tag_var.get().strip()
        t = LANG_DICT[self.lang]
        if new_tag:
            if new_tag in self.tags:
                messagebox.showerror("Error", t['error_tag_exists'])
            else:
                self.tags.append(new_tag)
                self.locked_tags[new_tag] = False # الحالة الافتراضية مفتوح
                self.new_tag_var.set("")
                self.refresh_tags_display()
                self.refresh_input_frame()

    def refresh_tags_display(self):
        t = LANG_DICT[self.lang]
        self.current_tags_lbl.config(text=f"{t['current_tags']} {', '.join(self.tags)}")

    def toggle_lock(self, tag):
        """دالة مسؤولة عن عكس حالة القفل وتغيير شكل الحقل"""
        is_locked = self.locked_tags[tag]
        self.locked_tags[tag] = not is_locked # عكس الحالة
        
        btn = self.lock_buttons[tag]
        entry = self.entry_widgets[tag]
        
        if self.locked_tags[tag]:
            btn.config(text="🔒")
            entry.config(state='readonly') # منع التعديل
        else:
            btn.config(text="🔓")
            entry.config(state='normal') # السماح بالتعديل

    def refresh_input_frame(self):
        for widget in self.dynamic_entries_frame.winfo_children():
            widget.destroy()
            
        self.entry_widgets.clear()
        self.lock_buttons.clear()
        
        for idx, tag in enumerate(self.tags):
            # اسم الـ Tag
            lbl = ttk.Label(self.dynamic_entries_frame, text=f"{tag}:")
            lbl.grid(row=0, column=idx*3, padx=2, pady=5)
            
            # حقل الإدخال
            entry = ttk.Entry(self.dynamic_entries_frame, width=15)
            entry.grid(row=0, column=idx*3 + 1, padx=2, pady=5)
            entry.bind('<Return>', lambda e: self.add_certificate())
            self.entry_widgets[tag] = entry
            
            # زر القفل
            is_locked = self.locked_tags.get(tag, False)
            btn_text = "🔒" if is_locked else "🔓"
            lock_btn = ttk.Button(self.dynamic_entries_frame, text=btn_text, width=3,
                                  command=lambda t=tag: self.toggle_lock(t))
            lock_btn.grid(row=0, column=idx*3 + 2, padx=2, pady=5)
            self.lock_buttons[tag] = lock_btn
            
            # تطبيق الحالة المحفوظة (في حال تغيير اللغة مثلاً)
            if is_locked:
                entry.config(state='readonly')
            
        if self.tags:
            # وضع المؤشر في أول حقل غير مقفول
            for tag in self.tags:
                if not self.locked_tags[tag]:
                    self.entry_widgets[tag].focus_set()
                    break

    def add_certificate(self):
        t = LANG_DICT[self.lang]
        cert_data = {}
        
        # قراءة البيانات (الدالة .get() تعمل حتى لو كان الحقل readonly)
        for tag, entry in self.entry_widgets.items():
            val = entry.get().strip()
            if not val:
                messagebox.showwarning("Warning", t['error_empty_fields'])
                return
            cert_data[tag] = val
            
        self.certificates.append(cert_data)
        
        # 🟢 مسح محتوى الحقول "الغير مقفلة فقط" 🟢
        first_unlocked_entry = None
        for tag, entry in self.entry_widgets.items():
            if not self.locked_tags[tag]:
                entry.delete(0, tk.END)
                if first_unlocked_entry is None:
                    first_unlocked_entry = entry
        
        # إعادة التركيز (Focus) لأول حقل تم تفريغه (غالباً سيكون اسم الطالب)
        if first_unlocked_entry:
            first_unlocked_entry.focus_set()
        
        self.render_all_certificates()

    def remove_certificate(self, index):
        del self.certificates[index]
        self.render_all_certificates()

    def render_all_certificates(self):
        for widget in self.scrollable_list.winfo_children():
            widget.destroy()
            
        t = LANG_DICT[self.lang]
        for index, cert_data in enumerate(self.certificates):
            row_frame = ttk.Frame(self.scrollable_list)
            row_frame.pack(fill='x', pady=2, padx=5)
            
            display_text = " | ".join([f"{k}: {v}" for k, v in cert_data.items()])
            ttk.Label(row_frame, text=display_text).pack(side='left', padx=5)
            
            del_btn = ttk.Button(row_frame, text=t['del_btn'], width=3,
                               command=lambda i=index: self.remove_certificate(i))
            del_btn.pack(side='right', padx=5)

    def status_update(self, msg):
        self.root.after(0, lambda: self.status_lbl.config(text=msg))

    def start_generation(self):
        t = LANG_DICT[self.lang]
        if not self.template_path_var.get():
            messagebox.showerror("Error", t['error_no_template'])
            return
        if not self.output_path_var.get():
            messagebox.showerror("Error", t['error_no_output'])
            return
        if not self.certificates:
            messagebox.showerror("Error", t['error_no_certs'])
            return
            
        self.generate_btn.config(state='disabled')
        self.status_update(t['msg_generating'])
        
        threading.Thread(target=self.generate_logic, daemon=True).start()

    def generate_logic(self):
        t = LANG_DICT[self.lang]
        template_path = self.template_path_var.get()
        output_folder = self.output_path_var.get()
        
        pythoncom.CoInitialize() 
        
        try:
            abs_template = os.path.abspath(template_path)
            abs_output = os.path.abspath(output_folder)
            
            word_app = win32com.client.DispatchEx("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
            
            total = len(self.certificates)
            for index, cert_data in enumerate(self.certificates, start=1):
                base_name = str(list(cert_data.values())[0]) 
                safe_name = "".join([c for c in base_name if c.isalnum() or c in ' -_']).strip()
                
                status_msg = f"Processing ({index}/{total}): {safe_name}" if self.lang == 'EN' else f"جاري معالجة ({index}/{total}): {safe_name}"
                self.status_update(status_msg)
                
                doc = DocxTemplate(abs_template)
                doc.render(cert_data)
                
                temp_docx_path = os.path.join(abs_output, f"temp_{safe_name}.docx")
                pdf_path = os.path.join(abs_output, f"Cert_{safe_name}.pdf")
                
                doc.save(temp_docx_path)
                
                word_doc = None
                try:
                    word_doc = word_app.Documents.Open(temp_docx_path)
                    word_doc.SaveAs(pdf_path, FileFormat=17) 
                except Exception as e:
                    print(f"Error PDF: {e}")
                finally:
                    if word_doc:
                        word_doc.Close(SaveChanges=False)
                
                if os.path.exists(temp_docx_path):
                    os.remove(temp_docx_path)
                    
            self.status_update(t['msg_done'])
            
        except Exception as e:
            self.status_update(f"Error: {e}")
        finally:
            word_app.Quit()
            pythoncom.CoUninitialize() 
            self.root.after(0, lambda: self.generate_btn.config(state='normal'))


if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateApp(root)
    root.mainloop()