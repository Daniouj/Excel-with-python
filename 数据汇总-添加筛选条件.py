import pandas as pd  # 导入pandas库，用于数据处理
import tkinter as tk  # 导入tkinter库，用于创建GUI界面
from tkinter import ttk, messagebox  # 从tkinter库导入ttk模块和messagebox，用于增强的GUI小部件和消息框
from tkinterdnd2 import DND_FILES, TkinterDnD  # 从tkinterdnd2库导入DND_FILES和TkinterDnD，用于拖放功能

def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)  # 尝试读取Excel文件
        return df  # 返回读取的数据框
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel file: {e}")  # 如果读取出错，显示错误消息框
        return None  # 返回None

def filter_and_summarize(df, columns_to_filter, column_to_summarize):
    try:
        filtered_df = df.drop_duplicates(subset=columns_to_filter)  # 根据指定列去重
        summary = filtered_df.groupby(column_to_summarize).size().reset_index(name="Count")  # 计数并生成汇总表
        return summary  # 返回汇总结果
    except KeyError as e:
        messagebox.showerror("Error", f"Column error: {e}")  # 如果列名错误，显示错误消息框
        return None  # 返回None
    except Exception as e:
        messagebox.showerror("Error", f"Error processing data: {e}")  # 处理数据过程中出现其他错误时显示错误消息框
        return None  # 返回None

def display_summary(summary, column_to_summarize):
    if summary is None:
        return  # 如果汇总结果为空，直接返回

    total_count = summary["Count"].sum()  # 计算Count列的总和

    root = tk.Tk()  # 创建主窗口
    root.title("Excel Summary")  # 设置窗口标题

    columns = ["Column", "Count"]

    tree = ttk.Treeview(root, columns=columns)  # 创建Treeview小部件
    tree.heading("#1", text=column_to_summarize)  # 设置第一列标题为汇总列名
    tree.heading("#2", text="Count")  # 设置第二列标题为“Count”

    for row in summary.itertuples():  # 遍历汇总结果
        values = [row[1], row.Count]
        tree.insert("", "end", values=values)  # 插入数据到Treeview

    # 插入总和行
    tree.insert("", "end", values=["总和", total_count])

    tree.pack()  # 打包Treeview小部件

    root.mainloop()  # 进入主循环，显示窗口

def get_columns_and_summarize(df):
    input_window = tk.Toplevel()  # 创建新的顶级窗口
    input_window.title("输入列名")  # 设置窗口标题为中文

    tk.Label(input_window, text="请输入要过滤的列名（用逗号分隔）:").pack(pady=5)  # 创建标签
    columns_entry = tk.Entry(input_window)  # 创建输入框用于输入列名
    columns_entry.pack(pady=5)

    tk.Label(input_window, text="请输入要汇总的列名:").pack(pady=5)  # 创建标签
    summarize_entry = tk.Entry(input_window)  # 创建输入框用于输入汇总列名
    summarize_entry.pack(pady=5)

    filters_frame = tk.Frame(input_window)
    filters_frame.pack(pady=5)

    filters = []

    def add_filter():
        filter_frame = tk.Frame(filters_frame)
        filter_frame.pack(pady=5)

        tk.Label(filter_frame, text="筛选列名:").pack(side="left")
        column_entry = tk.Entry(filter_frame)
        column_entry.pack(side="left")

        tk.Label(filter_frame, text="筛选值:").pack(side="left")
        value_entry = tk.Entry(filter_frame)
        value_entry.pack(side="left")

        filters.append((column_entry, value_entry))

    add_filter_button = tk.Button(input_window, text="添加筛选条件", command=add_filter)
    add_filter_button.pack(pady=5)

    def on_confirm():
        columns_to_filter = columns_entry.get().split(",")  # 获取过滤列名并分割为列表
        column_to_summarize = summarize_entry.get()  # 获取汇总列名

        if columns_to_filter and column_to_summarize:
            df_filtered = df.copy()
            for column_entry, value_entry in filters:
                column = column_entry.get()
                value = value_entry.get()
                if column and value:
                    df_filtered = df_filtered[df_filtered[column] == value]

            summary = filter_and_summarize(df_filtered, columns_to_filter, column_to_summarize)  # 过滤和汇总数据
            display_summary(summary, column_to_summarize)  # 显示汇总结果
        input_window.destroy()  # 关闭输入窗口

    confirm_button = tk.Button(input_window, text="确认", command=on_confirm)  # 创建确认按钮
    confirm_button.pack(pady=5)

def handle_drop(event):
    file_path = event.data  # 获取拖放的文件路径
    df = read_excel(file_path)  # 读取Excel文件
    if df is not None:
        get_columns_and_summarize(df)  # 弹出输入窗口获取用户输入并汇总数据

if __name__ == "__main__":
    root = TkinterDnD.Tk()  # 创建支持拖放功能的主窗口
    root.title("拖放Excel文件")  # 设置窗口标题为中文

    label = tk.Label(root, text="请将Excel文件拖放到此处:")  # 创建标签提示用户拖放文件
    label.pack()  # 打包标签

    root.drop_target_register(DND_FILES)  # 注册拖放目标
    root.dnd_bind('<<Drop>>', handle_drop)  # 绑定拖放事件处理函数

    root.mainloop()  # 进入主循环，显示窗口
