import tkinter as tk
from json import JSONDecodeError
from tkinter import filedialog

import ttkbootstrap as ttk
import ttkbootstrap.dialogs
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox

import json_tree
import pathlib


class Application(tk.Frame):

    def __init__(self, master):
        super().__init__(master)
        self.pack()

        self.master.minsize(width=700, height=240
        self.master.title("Json Tree Generator")

        self.create_widgets()

    def create_widgets(self):
        # 一行目
        container_1 = ttk.Frame(self)
        container_1.pack(fill=X, expand=YES, pady=5)
        # 二行目
        container_2 = ttk.Frame(self)
        container_2.pack(fill=X, expand=YES, pady=5)
        # 三行目
        container_3 = ttk.Frame(self)
        container_3.pack(fill=X, expand=YES, pady=5)
        # 四行目
        container_4 = ttk.Frame(self)
        container_4.pack(fill=X, expand=YES, pady=5)

        # jsonラベル
        self.json_label = ttk.Label(master=container_1, text="jsonファイル")
        self.json_label.pack(side=LEFT, padx=10, pady=10)

        # jsonファイルパス
        self.json_filepath_entry = ttk.Entry(master=container_1)
        self.json_filepath_entry.configure(width=50)
        self.json_filepath_entry.pack(side=LEFT, padx=10, pady=10)

        # jsonファイル参照ボタン
        self.json_filesearch_button = ttk.Button(master=container_1,
                                                 text='ファイル参照',
                                                 command=self.on_json_filepath_browse)
        self.json_filesearch_button.configure()
        self.json_filesearch_button.pack(padx=5, pady=10)

        # jsonラベル
        self.excel_label = ttk.Label(master=container_2, text="要素の終端に色を付ける.")
        self.excel_label.pack(side=LEFT, padx=10, pady=10)

        # Excelトグル
        self.excel_last_color_toggle_bool = tk.BooleanVar()
        self.excel_last_color_toggle_bool.set(False)  # 最初はチェックなし
        self.excel_last_color_toggle_checkbutton = ttk.Checkbutton(master=container_2,
                                                                   variable=self.excel_last_color_toggle_bool,
                                                                   bootstyle="round-toggle")
        self.excel_last_color_toggle_checkbutton.pack(side=LEFT, padx=10, pady=10)

        # Excel作成後オープントグル
        # ラベル
        self.excel_open_label = ttk.Label(master=container_3, text="作成後、Excelファイルを開く")
        self.excel_open_label.pack(side=LEFT, padx=10, pady=10)

        # トグル
        self.excel_open_bool = tk.BooleanVar()
        self.excel_open_bool.set(True)
        self.excel_open_toggle = ttk.Checkbutton(master=container_3,
                                                 variable=self.excel_open_bool,
                                                 bootstyle="round-toggle")
        self.excel_open_toggle.pack(side=LEFT, padx=10, pady=10)

        # 出力ボタン
        self.json_filesearch_button = ttk.Button(master=container_4, command=self.export_excel)
        self.json_filesearch_button.configure(text="出力")
        self.json_filesearch_button.pack(padx=15, pady=10)

    def on_json_filepath_browse(self):
        """
        ファイル参照
        :return:
        """
        f_type = [('JSONファイル', '*.json'), ('TEXTファイル', '*.txt')]
        path = filedialog.askopenfilename(title="ファイルを選択", filetype=f_type)
        if path:
            self.json_filepath_entry.delete(0, tk.END)
            self.json_filepath_entry.insert(0, path)

    def export_excel(self):
        """
        ファイル出力
        :return:
        """
        if len(self.json_filepath_entry.get()) == 0:
            Messagebox.show_warning('JSONファイルのパスを指定してください.', title=' ', parent=None)
            return

        f_type = [('Excel', '*.xlsx')]
        excel_path = filedialog.asksaveasfilename(filetype=f_type)

        if excel_path:
            if not excel_path.endswith('.xlsx'):
                excel_path += '.xlsx'

            try:
                json_tree.main(self.json_filepath_entry.get(),
                               excel_path,
                               self.excel_last_color_toggle_bool.get(),
                               self.excel_open_bool.get())
            except FileNotFoundError:
                Messagebox.show_error('JSONファイルが見つかりません', title='FileNotFoundError', parent=None)
            except JSONDecodeError:
                Messagebox.show_error('JSONファイルとして読み取ることができませんでした', title='JSONDecodeError', parent=None)


def start_point():
    root = ttk.Window(themename="litera")  # ダークテーマとか
    app = Application(master=root)  # Inherit

    app.mainloop()


if __name__ == "__main__":
    start_point()
