########################################################################################################
# Projet : PDFtoDocx                                                                                   #
# Auteur : Soradev                                                                                     #
# Version : 1.0.0                                                                                      #
########################################################################################################
# Description :                                                                                        #
#   Convert PDF to DOCX                                                                                #
########################################################################################################
# For any questions or contributions, please contact the author at sora.dev.pro@gmail.com              #
########################################################################################################


import fitz 
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import Scrollbar, Canvas
import threading
from PIL import Image, ImageTk
import os
import sys
import webbrowser 
from Convert import Converter  

class PDFViewer(tk.Tk):
    def __init__(self, pdf_path=None):
        super().__init__()
        self.title("PDF Viewer")
        self.geometry("1200x600")
        self.resizable(False, False) 

        self.base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)

        self.background_image = tk.PhotoImage(file=os.path.join(self.base_path, "resource/background600.png"))
        self.background_label = tk.Label(self, image=self.background_image)
        self.background_label.place(relwidth=1, relheight=1)

        self.open_image = tk.PhotoImage(file=os.path.join(self.base_path, "resource/btn/open.png"))
        self.convert_image = tk.PhotoImage(file=os.path.join(self.base_path, "resource/btn/word.png"))
        self.print_image = tk.PhotoImage(file=os.path.join(self.base_path, "resource/btn/print.png"))
        self.zoom_in_image = tk.PhotoImage(file=os.path.join(self.base_path, "resource/btn/zoomav.png"))
        self.zoom_out_image = tk.PhotoImage(file=os.path.join(self.base_path, "resource/btn/zoomar.png"))

        button_frame = tk.Frame(self, highlightthickness=0)
        button_frame.pack(side=tk.TOP, pady=10, padx=20)

        self.open_canvas = Canvas(button_frame, width=173, height=28, bd=0, highlightthickness=0)
        self.open_canvas.pack(side=tk.LEFT, padx=20)
        self.open_canvas.create_image(0, 0, anchor="nw", image=self.open_image)
        self.open_canvas.bind("<Button-1>", lambda e: self.open_pdf())
        self.open_canvas.config(cursor="hand2")

        self.convert_canvas = Canvas(button_frame, width=188, height=28, bd=0, highlightthickness=0)
        self.convert_canvas.pack(side=tk.LEFT, padx=20)
        self.convert_canvas.create_image(0, 0, anchor="nw", image=self.convert_image)
        self.convert_canvas.bind("<Button-1>", lambda e: self.start_conversion_thread())
        self.convert_canvas.config(cursor="hand2")

        self.print_canvas = Canvas(button_frame, width=115, height=28, bd=0, highlightthickness=0)
        self.print_canvas.pack(side=tk.LEFT, padx=20)
        self.print_canvas.create_image(0, 0, anchor="nw", image=self.print_image)
        self.print_canvas.bind("<Button-1>", lambda e: self.print_pdf())
        self.print_canvas.config(cursor="hand2")

        self.canvas = Canvas(self, bg="#f0f0f0")
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        nav_frame = tk.Frame(self, highlightthickness=0)
        nav_frame.pack(side=tk.BOTTOM, pady=10)

        self.prev_button = tk.Button(nav_frame, text="Précédent", command=self.prev_page, state=tk.DISABLED)
        self.prev_button.pack(side=tk.LEFT, padx=10)

        self.page_label = tk.Label(nav_frame, text="Page 1 / 1")
        self.page_label.pack(side=tk.LEFT)

        self.next_button = tk.Button(nav_frame, text="Suivant", command=self.next_page, state=tk.DISABLED)
        self.next_button.pack(side=tk.LEFT, padx=10)

        self.zoom_out_canvas = Canvas(self, width=98, height=28)
        self.zoom_out_canvas.create_image(0, 0, anchor="nw", image=self.zoom_out_image)
        self.zoom_out_canvas.bind("<Button-1>", lambda e: self.zoom_out())
        self.zoom_out_canvas.config(cursor="hand2")
        self.zoom_out_canvas.place(x=625, y=486) 

        self.zoom_in_canvas = Canvas(self, width=98, height=28, bd=0, highlightthickness=0)
        self.zoom_in_canvas.create_image(0, 0, anchor="nw", image=self.zoom_in_image)
        self.zoom_in_canvas.bind("<Button-1>", lambda e: self.zoom_in())
        self.zoom_in_canvas.config(cursor="hand2")
        self.zoom_in_canvas.place(x=478, y=486) 

        self.scroll_x = Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.scroll_y = Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.configure(xscrollcommand=self.scroll_x.set, yscrollcommand=self.scroll_y.set)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.pdf_path = pdf_path
        self.images = []  
        self.zoom_level = 1.0  
        self.current_page = 0  
        self.total_pages = 0  

        if self.pdf_path:
            self.load_pdf(self.pdf_path)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def open_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.pdf_path = file_path
            self.load_pdf(self.pdf_path)
            self.update_title()

    def load_pdf(self, pdf_path):
        self.pdf_document = fitz.open(pdf_path)
        self.total_pages = self.pdf_document.page_count

        self.canvas.delete("all")

        self.images.clear() 
        self.pages = [self.pdf_document.load_page(i) for i in range(self.pdf_document.page_count)]
        self.display_page(self.current_page)

        self.update_navigation_buttons()

    def display_page(self, page_num):
        self.current_page = page_num
        page = self.pages[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level))
        img_data = pix.tobytes("ppm")

        image = tk.PhotoImage(data=img_data)
        self.images = [image]  

        canvas_width = self.canvas.winfo_width()
        image_width = pix.width
        x_offset = max((canvas_width - image_width) // 2, 0)

        self.canvas.delete("all") 
        self.canvas.create_image(x_offset, 0, anchor="nw", image=image)
        self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))

        self.page_label.config(text=f"Page {self.current_page + 1} / {self.total_pages}")

    def prev_page(self):
        if self.current_page > 0:
            self.display_page(self.current_page - 1)
        self.update_navigation_buttons()

    def next_page(self):
        if self.current_page < self.total_pages - 1:
            self.display_page(self.current_page + 1)
        self.update_navigation_buttons()

    def update_navigation_buttons(self):
        """Mettre à jour l'état des boutons de navigation."""
        self.prev_button.config(state=tk.NORMAL if self.current_page > 0 else tk.DISABLED)
        self.next_button.config(state=tk.NORMAL if self.current_page < self.total_pages - 1 else tk.DISABLED)

    def zoom_in(self):
        self.zoom_level *= 1.2 
        self.display_page(self.current_page)

    def zoom_out(self):
        self.zoom_level /= 1.2
        self.display_page(self.current_page)

    def start_conversion_thread(self):
        conversion_thread = threading.Thread(target=self.convert_pdf)
        conversion_thread.start()

    def convert_pdf(self):
        output_path = self.pdf_path.replace(".pdf", ".docx")
        converter = Converter(pdf_file=self.pdf_path)
        converter.convert(output_file=output_path)
        converter.close()
        messagebox.showinfo("Conversion terminée", f"Le fichier a été converti en {output_path}")

    def print_pdf(self):
        print("Bouton imprimer cliqué !")

    def update_title(self):
        """Mettre à jour le titre de la fenêtre avec le nom du fichier PDF."""
        file_name = os.path.basename(self.pdf_path)
        self.title(f"PDF Viewer - {file_name}")

def main():
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else None
    root = PDFViewer(pdf_path=pdf_path)
    root.mainloop()

if __name__ == "__main__":
    main()
