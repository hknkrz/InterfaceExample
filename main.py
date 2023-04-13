import os.path

import customtkinter
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
from tkinter import ttk
import textwrap

LOGS_PATH = 'logs'
COLUMN_NAMES = dict()

customtkinter.set_default_color_theme("blue")


def wrap(string, lenght=22):
    return '\n'.join(textwrap.wrap(string, lenght))


class App(customtkinter.CTk):
    APP_NAME = "Sample text"
    WIDTH = 800
    HEIGHT = 600

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # self.attributes('-fullscreen', True)

        self.data = None
        self.graph_data = None
        self.style = ttk.Style()
        self.style.configure("Treeview", font=('Helvetica', 18), rowheight=60, wordbreak=True)
        self.style.configure("Treeview.Heading", font=('Helvetica', 20))
        self.title(App.APP_NAME)
        self.geometry(str(App.WIDTH) + "x" + str(App.HEIGHT))
        self.minsize(App.WIDTH, App.HEIGHT)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind("<Command-q>", self.on_closing)
        self.bind("<Command-w>", self.on_closing)
        self.createcommand('tk::mac::Quit', self.on_closing)

        self.marker_list = []

        # ============ create two CTkFrames ============

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame_left = customtkinter.CTkFrame(master=self, width=150, corner_radius=0, fg_color=None)
        self.frame_left.grid(row=0, column=0, padx=0, pady=0, sticky="nsew")

        self.frame_right = customtkinter.CTkFrame(master=self, corner_radius=0)
        self.frame_right.grid(row=0, column=1, rowspan=1, pady=0, padx=0, sticky="nsew")

        # ============ frame_left ============

        self.frame_left.grid_rowconfigure(2, weight=1)

        self.button_1 = customtkinter.CTkButton(master=self.frame_left,
                                                text="Рассчитать",
                                                command=self.load_excel_file)
        self.button_1.grid(pady=(20, 0), padx=(20, 20), row=0, column=0)

        self.button_2 = customtkinter.CTkButton(master=self.frame_left,
                                                text="График",
                                                command=self.display_graph_event)
        self.button_2.grid(pady=(20, 0), padx=(20, 20), row=1, column=0)

        self.map_label = customtkinter.CTkLabel(self.frame_left, text="", anchor="w")
        self.map_label.grid(row=3, column=0, padx=(20, 20), pady=(20, 0))
        self.table_button = customtkinter.CTkButton(master=self.frame_left,
                                                    text="Таблица",
                                                    command=self.display_table)
        self.table_button.grid(row=4, column=0, padx=(20, 20), pady=(10, 0))

        self.appearance_mode_label = customtkinter.CTkLabel(self.frame_left, text="Цветовая гамма:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=(20, 20), pady=(20, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.frame_left,
                                                                       values=["Light", "Dark"],
                                                                       command=self.change_appearance_mode)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=(20, 20), pady=(10, 20))

        # ============ frame_right ============

        self.frame_right.grid_rowconfigure(2, weight=1)
        self.frame_right.grid_rowconfigure(0, weight=0)
        self.frame_right.grid_rowconfigure(3, weight=1)
        self.frame_right.grid_columnconfigure(0, weight=1)
        self.frame_right.grid_columnconfigure(1, weight=0)
        self.frame_right.grid_columnconfigure(2, weight=0)

        self.entry = customtkinter.CTkEntry(master=self.frame_right,
                                            placeholder_text="Введите путь к файлу excel с данными", width=300)
        self.entry.grid(row=0, column=0, sticky="we", padx=(12, 0), pady=12)
        self.error_frame = tk.Frame(self.frame_right)
        self.error_frame.pack_forget()
        self.error_label = customtkinter.CTkLabel(self.error_frame, text="")
        self.error_label.pack(side=tk.TOP)

        self.entry.bind("<Return>", self.search_event)

        self.button_5 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Поиск",
                                                width=90,
                                                command=self.search_event)
        self.button_5.grid(row=0, column=1, sticky="w", padx=(12, 0), pady=12)

        self.appearance_mode_optionemenu.set("Light")

    def update_data(self, data_=None, graph_data_=None):
        self.data = data_
        self.graph_data = graph_data_

    def search_event(self, event=None):
        self.error_label.configure(text="")
        self.error_frame.grid_remove()
        self.data = None
        file_path = filedialog.askopenfilename(
            title="Выберите файл",
            filetypes=[("Excel files", "*.xlsx")]
        )
        self.entry.delete(0, len(self.entry.get()))
        self.entry.insert(0, file_path)

    def load_excel_file(self):
        self.error_label.configure(text="")
        self.error_frame.grid_remove()
        filename = self.entry.get()
        try:
            df = pd.read_excel(filename)
            self.error_label.configure(text="")
            self.error_frame.grid_remove()
        except:
            self.error_label.configure(text="Неверный путь к файлу")
            self.error_frame.grid(row=1, column=0, sticky="we", padx=(12, 0), pady=12)
            return
        try:
            df.columns = ['col{}'.format(i + 1) for i in range(len(list(df.columns)))]
            df = df[4:]
            invalid_rows = set()
            for column in set(df.columns).difference({'col1'}):
                invalid_rows.update(set(df[pd.to_numeric(df[column], errors='coerce').isna()]['col1'].values))

            if len(invalid_rows) != 0:
                with open(os.path.join(LOGS_PATH, 'logs.txt'), 'w') as f:
                    f.write('\n'.join(sorted(invalid_rows)))
                self.error_label.configure(text="Некорректные данные. Информация о некорректных строках в папке logs")
                self.error_frame.grid(row=1, column=0, sticky="we", padx=(12, 0), pady=12)
                return
            self.data = [[df['col4'].mean(), df['col6'].mean(), df['col8'].mean()]
                , [df['col4'].median(), df['col6'].median(), df['col8'].median()]
                , [df['col4'].max(), df['col6'].max(), df['col8'].max()]]
            self.error_label.configure(text="Коэффициенты рассчитаны")
            self.error_frame.grid(row=1, column=0, sticky="we", padx=(12, 0), pady=12)
        except:
            self.error_label.configure(text="Неверный формат файла")
            self.error_frame.grid(row=1, column=0, sticky="we", padx=(12, 0), pady=12)

        print(123)

    def set_marker_event(self):
        pass

    def display_graph_event(self):
        graph = customtkinter.CTkImage(Image.open("placeholder.jpg"), size=(600, 350))
        graph_label = customtkinter.CTkLabel(self.frame_right, image=graph, text="", padx=12, pady=12)
        graph_label.grid(row=3, column=0)

    def change_appearance_mode(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def display_table(self):
        if self.data is None:
            self.error_label.configure(text="Сначала загрузите файл с данными и рассчитайте коэффициенты")
            self.error_frame.grid(row=1, column=0, sticky="we", padx=(12, 0), pady=12)
            return

        tree = ttk.Treeview(self.frame_right)
        tree["columns"] = ("row", "row_1", "row_2", "row_3")
        tree.column("#0", width=0)
        tree.column("row", anchor='w', width=300)
        tree.column("row_1", anchor='center', width=300)
        tree.column("row_2", anchor='center', width=300)
        tree.column("row_3", anchor='center', width=300)
        tree.configure(height=4)

        tree.heading("#0", text='', anchor='w')
        tree.heading("row", text="Ряд", anchor='w')
        tree.heading("row_1", text="Ось-4,5 м ", anchor='center')
        tree.heading("row_2", text="Ось", anchor='center')
        tree.heading("row_3", text="Ось+4,5 м ", anchor='center')

        # Insert data
        tree.insert(parent='', index='end', iid=0, text='',
                    values=(wrap('Коэффициент мощности спектра С '), self.data[0][0], self.data[0][1], self.data[0][2]))
        tree.insert(parent='', index='end', iid=1, text='',
                    values=('Показатель степени k ', self.data[1][0], self.data[1][1], self.data[1][2]))
        tree.insert(parent='', index='end', iid=2, text='',
                    values=('Критерий ровности R ', self.data[2][0], self.data[2][1], self.data[2][2]))

        tree.grid(row=2, column=0)

    def on_closing(self, event=0):
        self.destroy()

    def start(self):
        self.mainloop()


if __name__ == "__main__":
    app = App()
    app.start()
