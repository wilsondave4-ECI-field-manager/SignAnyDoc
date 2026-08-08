import os
import tempfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import fitz
import doc_sign as core


def open_help(self):
    try:
        help_path = core.resource_path('README.md')
        text = help_path.read_text(encoding='utf-8')
    except Exception:
        text = 'Doc Sign help file could not be loaded.'
    win = tk.Toplevel(self.root)
    win.title('Doc Sign - Help / How to Use')
    win.geometry('780x700')
    win.minsize(620, 500)
    outer = ttk.Frame(win, padding=12)
    outer.pack(fill='both', expand=True)
    ttk.Label(outer, text='Doc Sign - Help / How to Use', font=('Segoe UI', 16, 'bold')).pack(anchor='w')
    ttk.Label(outer, text='By Dave Wilson', font=('Segoe UI', 10)).pack(anchor='w', pady=(0, 10))
    box = ScrolledText(outer, wrap='word', font=('Segoe UI', 10))
    box.pack(fill='both', expand=True)
    box.insert('1.0', text)
    box.configure(state='disabled')
    ttk.Button(outer, text='Close', command=win.destroy).pack(anchor='e', pady=(10, 0))


def build_ui_v2(self):
    top = ttk.Frame(self.root, padding=12)
    top.pack(fill='x')
    title_box = ttk.Frame(top)
    title_box.pack(side='left')
    ttk.Label(title_box, text='Doc Sign - Local Document Signing', font=('Segoe UI', 18, 'bold')).pack(anchor='w')
    ttk.Label(title_box, text='By Dave Wilson', font=('Segoe UI', 10)).pack(anchor='w')

    ttk.Button(top, text='Open PDF', command=self.open_pdf).pack(side='right', padx=4)
    ttk.Button(top, text='Open Word', command=self.open_word).pack(side='right', padx=4)
    self.close_doc_btn = ttk.Button(top, text='Close Document', command=self.close_document)
    self.close_doc_btn.pack(side='right', padx=4)
    self.png_btn = ttk.Button(top, text='Save Signature PNG', command=self.save_signature_png, state='disabled')
    self.png_btn.pack(side='right', padx=4)
    self.help_btn = ttk.Button(top, text='Help / How to Use', command=self.open_help)
    self.help_btn.pack(side='right', padx=4)

    body = ttk.Panedwindow(self.root, orient='horizontal')
    body.pack(fill='both', expand=True, padx=12, pady=(0, 12))
    left = ttk.Frame(body, padding=8)
    right_host = ttk.Frame(body, width=350)
    body.add(left, weight=4)
    body.add(right_host, weight=1)

    toolbar = ttk.Frame(left)
    toolbar.pack(fill='x')
    self.file_label = ttk.Label(toolbar, text='No document loaded')
    self.file_label.pack(side='left')
    self.page_label = ttk.Label(toolbar, text='')
    self.page_label.pack(side='right')
    ttk.Button(toolbar, text='<', width=4, command=self.prev_page).pack(side='right', padx=2)
    ttk.Button(toolbar, text='>', width=4, command=self.next_page).pack(side='right', padx=2)

    self.canvas = tk.Canvas(left, bg='#e8edf2', highlightthickness=1, highlightbackground='#ccd6de')
    self.canvas.pack(fill='both', expand=True, pady=8)
    self.canvas.bind('<Configure>', lambda e: self.render_current())
    self.canvas.bind('<Button-1>', self.place_signature)

    actions = ttk.Frame(left)
    actions.pack(fill='x')
    ttk.Label(actions, text='Signature size').pack(side='left')
    ttk.Button(actions, text='-', width=4, command=lambda: self.resize_sig(-.015)).pack(side='left', padx=3)
    ttk.Button(actions, text='+', width=4, command=lambda: self.resize_sig(.015)).pack(side='left')
    self.save_btn = ttk.Button(actions, text='Save signed document', command=self.save_signed, state='disabled')
    self.save_btn.pack(side='right')
    self.confirm_btn = ttk.Button(actions, text='Confirm on this page & Next', command=self.confirm_signature, state='disabled')
    self.confirm_btn.pack(side='right', padx=8)

    # Scrollable right-side panel. This prevents the received signature preview from
    # hiding the Clear Signature / Security controls when Doc Sign is not maximized.
    right_canvas = tk.Canvas(right_host, highlightthickness=0, borderwidth=0)
    right_scroll = ttk.Scrollbar(right_host, orient='vertical', command=right_canvas.yview)
    right_canvas.configure(yscrollcommand=right_scroll.set)
    right_scroll.pack(side='right', fill='y')
    right_canvas.pack(side='left', fill='both', expand=True)
    right = ttk.Frame(right_canvas, padding=(8, 0, 8, 8))
    right_window = right_canvas.create_window((0, 0), window=right, anchor='nw')

    def update_scrollregion(event=None):
        right_canvas.configure(scrollregion=right_canvas.bbox('all'))

    def fit_right_width(event):
        right_canvas.itemconfigure(right_window, width=max(event.width, 250))

    def mousewheel(event):
        right_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    right.bind('<Configure>', update_scrollregion)
    right_canvas.bind('<Configure>', fit_right_width)
    right_canvas.bind('<Enter>', lambda e: right_canvas.bind_all('<MouseWheel>', mousewheel))
    right_canvas.bind('<Leave>', lambda e: right_canvas.unbind_all('<MouseWheel>'))
    self.right_scroll_canvas = right_canvas

    pair = ttk.LabelFrame(right, text='Phone pairing', padding=14)
    pair.pack(fill='x')
    self.code_label = ttk.Label(pair, text='', font=('Segoe UI', 22, 'bold'))
    self.code_label.pack()
    self.pin_label = ttk.Label(pair, text='', font=('Segoe UI', 18, 'bold'))
    self.pin_label.pack(pady=(0, 8))
    self.qr_label = ttk.Label(pair)
    self.qr_label.pack(pady=6)
    self.url_box = tk.Text(pair, height=4, width=34, wrap='word')
    self.url_box.pack(fill='x')
    self.url_box.configure(state='disabled')
    ttk.Button(pair, text='Open phone page on this PC', command=self.open_phone_local).pack(fill='x', pady=(8, 0))
    ttk.Button(pair, text='New session', command=self.new_session).pack(fill='x', pady=5)
    self.pair_status = ttk.Label(pair, text='Waiting for phone')
    self.pair_status.pack(pady=5)

    sig = ttk.LabelFrame(right, text='Received signature', padding=14)
    sig.pack(fill='x', pady=12)
    self.sig_preview = ttk.Label(sig, text='No signature received')
    self.sig_preview.pack(pady=8)
    ttk.Button(sig, text='Save Signature PNG', command=self.save_signature_png).pack(fill='x', pady=(0, 6))
    ttk.Button(sig, text='Clear signature', command=self.clear_signature).pack(fill='x')

    note = ttk.LabelFrame(right, text='Security', padding=12)
    note.pack(fill='x')
    ttk.Label(note, text='Local network only\nSession code + PIN\nOriginal document is not overwritten', justify='left').pack(anchor='w')

    self.status = ttk.Label(self.root, text='Ready.', anchor='w', padding=(12, 4))
    self.status.pack(fill='x')


def close_preview_v2(self):
    try:
        if self.doc:
            self.doc.close()
    except Exception:
        pass
    self.doc = None
    old_preview = getattr(self, 'preview_pdf_path', None)
    self.preview_pdf_path = None
    if old_preview:
        try:
            Path(old_preview).unlink(missing_ok=True)
        except Exception:
            pass


def close_document(self):
    if not self.doc and not self.path:
        return
    if self.confirmed_signatures:
        if not messagebox.askyesno('Close document', 'This document has confirmed signature placements that have not been saved.\n\nClose it and discard those placements?'):
            return
    self.close_preview()
    self.path = None
    self.doc_type = None
    self.page_index = 0
    self.confirmed_signatures = []
    self.file_label.config(text='No document loaded')
    self.page_label.config(text='')
    self.canvas.delete('all')
    self.confirm_btn.config(state='disabled')
    self.save_btn.config(state='disabled')
    self.status.config(text='Document closed. Ready to open another PDF or Word document.')


def open_word_v2(self):
    if not core.WORD_COM_AVAILABLE:
        messagebox.showerror('Open Word', 'Microsoft Word desktop integration is not available on this computer.')
        return
    p = filedialog.askopenfilename(filetypes=[('Word documents', '*.doc;*.docx;*.docm'), ('All files', '*.*')])
    if not p:
        return
    self.close_preview()
    self.status.config(text='Opening Word document inside Doc Sign...')
    self.root.update_idletasks()
    word = None
    wdoc = None
    preview = None
    try:
        core.pythoncom.CoInitialize()
        temp_dir = Path(tempfile.gettempdir()) / 'DocSign'
        temp_dir.mkdir(parents=True, exist_ok=True)
        fd, preview_name = tempfile.mkstemp(prefix='word_preview_', suffix='.pdf', dir=str(temp_dir))
        os.close(fd)
        preview = Path(preview_name)
        preview.unlink(missing_ok=True)
        word = core.win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        wdoc = word.Documents.Open(str(Path(p).resolve()), ReadOnly=True, AddToRecentFiles=False)
        wdoc.ExportAsFixedFormat(str(preview), core.WD_EXPORT_FORMAT_PDF)
        wdoc.Close(False)
        wdoc = None
        word.Quit()
        word = None
        self.doc = fitz.open(str(preview))
        self.preview_pdf_path = preview
        self.path = Path(p)
        self.doc_type = 'word'
        self.page_index = 0
        self.confirmed_signatures = []
        self.file_label.config(text=f'{self.path.name}  [Word preview]')
        self.confirm_btn.config(state='normal' if self.signature_pil else 'disabled')
        self.save_btn.config(state='disabled')
        self.status.config(text='Word opened inside Doc Sign. Place and confirm signatures exactly where required.')
        self.render_current()
    except Exception as e:
        try:
            if wdoc is not None:
                wdoc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        if preview:
            try:
                preview.unlink(missing_ok=True)
            except Exception:
                pass
        messagebox.showerror('Open Word', 'Could not render the Word document inside Doc Sign.\n\n' + str(e))
    finally:
        try:
            core.pythoncom.CoUninitialize()
        except Exception:
            pass


core.App.build_ui = build_ui_v2
core.App.close_preview = close_preview_v2
core.App.close_document = close_document
core.App.open_word = open_word_v2
core.App.open_help = open_help


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use('vista')
    except Exception:
        pass
    app = core.App(root)
    try:
        app.server.daemon_threads = True
    except Exception:
        pass

    shutting_down = {'value': False}

    def shutdown():
        if shutting_down['value']:
            return
        shutting_down['value'] = True
        try:
            app.close_preview()
        except Exception:
            pass
        try:
            if app.server:
                app.server.shutdown()
                app.server.server_close()
        except Exception:
            pass
        try:
            root.destroy()
        except Exception:
            pass
        os._exit(0)

    root.protocol('WM_DELETE_WINDOW', shutdown)
    root.mainloop()


if __name__ == '__main__':
    main()
