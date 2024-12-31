import tkinter as tk
from tkinter import scrolledtext


class TextCollectorGUI:
    def __init__(self, master):
        self.master = master
        master.title("多行文字输入")

        self.text_area = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=40, height=10)
        self.text_area.pack(padx=10, pady=10)

        self.button = tk.Button(master, text="获取输入内容", command=self.get_text_content)
        self.button.pack(pady=5)

        self.result_text = tk.StringVar()
        self.result_label = tk.Label(master, textvariable=self.result_text)
        self.result_label.pack(pady=5)

    def get_text_content(self):
        input_text = self.text_area.get("1.0", tk.END)
        self.result_text.set("你输入的内容如下：\n" + input_text)


if __name__ == "__main__":
    root = tk.Tk()
    app = TextCollectorGUI(root)
    root.mainloop()