import tempfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz
import doc_sign as core


# Keep the proven signing engine, but replace document lifecycle/Word preview handling.
_original_build_ui = core.App.build_ui
_original_close_preview = core.App.close_preview


def build_ui_v2(self):
    _original_build_ui(self)
    # The first root child is the top toolbar created by the base UI.
    try:
        top = self.root.winfo_children()[0]
        self.close_doc_btn = ttk.Button(top, text='Close Document', command=self.close_document)
        self.close_doc_btn.pack(side='right', padx=8)
    except Exception:
        pass


def close_preview_v2(self):
    # Close PyMuPDF first. On Windows an open PDF handle prevents deletion/re-use.
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
            # A stale temp preview is harmless because every new Word preview gets a unique filename.
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

    # Release the previous preview before starting Word. This fixes WinError 32.
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
        # Never reuse word_preview.pdf: Windows/PyMuPDF may still hold the previous file briefly.
        fd, preview_name = tempfile.mkstemp(prefix='word_preview_', suffix='.pdf', dir=str(temp_dir))
        import os
        os.close(fd)
        preview = Path(preview_name)
        preview.unlink(missing_ok=True)  # Word requires a destination it can create itself.

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


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use('vista')
    except Exception:
        pass
    app = core.App(root)

    def shutdown():
        try:
            app.close_preview()
        except Exception:
            pass
        try:
            if app.server:
                app.server.shutdown()
        except Exception:
            pass
        root.destroy()

    root.protocol('WM_DELETE_WINDOW', shutdown)
    root.mainloop()


if __name__ == '__main__':
    main()
