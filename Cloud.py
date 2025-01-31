import tkinter as tk # 导入tkinter库用于创建GUI
from tkinter import filedialog, messagebox, ttk # 导入tkinter特定模块以增强GUI功能
import os # 导入操作系统接口模块
import jieba # 导入jieba库用于分词
from wordcloud import WordCloud # 导入wordcloud库用于生成词云
from PIL import Image, ImageTk # 导入PIL库中的Image和ImageTk用于图像处理
from docx import Document # 导入docx库用于处理Word文档
import win32com.client as win32  # 导入win32com库用于与Windows应用程序通信
import numpy as np # 导入numpy库用于数值计算
import sys # 导入sys模块用于系统相关的操作

#获取相对路径
def get_resource_path(relative_path):
    # 如果使用PyInstaller打包，则从_MEIPASS获取路径
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    # 否则从当前工作目录获取路径
    return os.path.join(os.path.abspath("."), relative_path)

# 词云生成器类
class WordCloudGenerator:
    def __init__(self, root):
        # 初始化主窗口
        self.root = root
        self.root.title("词云生成系统(@hsl)")
        self.root.iconbitmap(default=get_resource_path("images/cloud.ico"))
        width = 400
        heigh = 200
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        self.root.geometry('%dx%d+%d+%d'%(width, heigh, (screenwidth-width)/2, (screenheight-heigh)/2))

        # 初始化变量
        self.directory_path = ""  # 文档目录路径
        self.threshold = 1  # 词频阈值
        self.layout = "默认"  # 词云布局
        self.freq_dict = {}  # 词频字典
        font_style = ("images/Serif.ttf", 12) 
         #创建并放置标签、输入框和按钮等GUI组件
        frame_dir = tk.Frame(root)
        frame_dir.pack(pady=10, fill=tk.X)
        frame_dir.grid_columnconfigure(0, weight=1)
        frame_dir.grid_columnconfigure(1, weight=3)
        frame_dir.grid_columnconfigure(2, weight=1)
        self.label_dir = tk.Label(frame_dir, text="文档目录:", font=font_style)
        self.label_dir.grid(row=0, column=0, padx=(5, 0), sticky=tk.W)
        self.entry_dir = tk.Entry(frame_dir, width=30, font=font_style)
        self.entry_dir.grid(row=0, column=1, padx=(0, 5), sticky=tk.EW)
        self.button_browse = tk.Button(frame_dir, text="浏览", command=self.browse_directory, width=0, font=font_style)
        self.button_browse.grid(row=0, column=2, padx=(0, 5), sticky=tk.EW)  # 宽度设为0，使按钮不可见
        
        # 创建词频阈值和词云布局选择框
        frame_threshold_layout = tk.Frame(root)
        frame_threshold_layout.pack(pady=10, fill=tk.X)
        frame_threshold_layout.grid_columnconfigure(0, weight=1)
        frame_threshold_layout.grid_columnconfigure(1, weight=1)
        frame_threshold_layout.grid_columnconfigure(2, weight=1)
        frame_threshold_layout.grid_columnconfigure(3, weight=1)
        
        self.label_threshold = tk.Label(frame_threshold_layout, text="词频阈值:", font=font_style)
        self.label_threshold.grid(row=0, column=0, padx=(5, 0), pady=0, sticky=tk.W)
        self.entry_threshold = tk.Entry(frame_threshold_layout, width=10, font=font_style)
        self.entry_threshold.grid(row=0, column=1, padx=(0, 5), pady=0, sticky=tk.EW)
        
        self.label_layout = tk.Label(frame_threshold_layout, text="词云布局:", font=font_style)
        self.label_layout.grid(row=0, column=2, padx=(5, 0), pady=0, sticky=tk.W)
        self.combo_layout = ttk.Combobox(frame_threshold_layout, values=["长方形","圆形","爱心", "五角星", "气泡","奇异","四角星"], width=6, font=font_style)
        self.combo_layout.current(0)
        self.combo_layout.grid(row=0, column=3, padx=(0, 55), pady=0, sticky=tk.EW)
        
        self.button_generate = tk.Button(root, text="生成词云", command=self.generate_wordcloud, font=font_style)
        self.button_generate.pack(pady=20)
        
        # 添加一个标签用于显示生成中的状态信息
        self.status_label = tk.Label(root, text="", font=("images/Serif.ttf", 14))
        self.status_label.pack(pady=10)

    # 浏览目录按钮点击事件
    def browse_directory(self):
        """
        弹出文件夹选择对话框，并获取选择的目录路径。
        如果选择了目录，则将其路径插入输入框。
        """
        self.directory_path = filedialog.askdirectory()  
        if self.directory_path:
            self.entry_dir.delete(0, tk.END)  
            self.entry_dir.insert(0, self.directory_path)  
    # 读取.doc文件内容
    def read_doc_file(self, filepath):
        """
        使用Microsoft Word读取.doc文件内容。
        """
        word = win32.gencache.EnsureDispatch('Word.Application')  
        word.Visible = False  
        doc = word.Documents.Open(filepath)  
        content = doc.Content.Text  
        doc.Close(False)  
        word.Quit()  
        return content  
    
    # 读取.docx文件内容
    def read_docx_file(self, filepath):
        """
        使用python-docx库读取.docx文件内容。
        """
        doc = Document(filepath)  
        fullText = []
        for para in doc.paragraphs:  
            fullText.append(para.text)  
        return '\n'.join(fullText)  
    
    def generate_wordcloud(self):
        """
        根据用户设置生成词云。
        """
        try:
            self.threshold = int(self.entry_threshold.get())  # 获取词频阈值
        except ValueError:
            messagebox.showerror("错误", "请输入有效的整数作为词频阈值")  
            return
        self.layout = self.combo_layout.get()  # 获取词云布局
        self.status_label.config(text="生成中...")  # 显示生成中的状态
        self.root.update_idletasks()  
        
        self.all_text = ""
        for filename in os.listdir(self.directory_path):  
            filepath = os.path.join(self.directory_path, filename)
            if filename.endswith(".doc"):
                self.all_text += self.read_doc_file(filepath)  
            elif filename.endswith(".docx"):
                self.all_text += self.read_docx_file(filepath)  
        
        words = jieba.lcut(self.all_text)  # 使用结巴分词进行分词
        self.freq_dict = {}
        for word in words:  
            if len(word.strip()) > 1:  
                self.freq_dict[word] = self.freq_dict.get(word, 0) + 1  # 统计词频
        

        self.update_wordcloud_display()  # 更新词云显示
        self.open_edit_window()  # 打开编辑窗口
        
        self.status_label.config(text="")  # 清空状态信息
    
    def update_wordcloud_display(self):
        """
        更新主窗口中的词云显示。
        """
        #不同形状的词云
        layout=self.layout;
        mask = None
        if layout == "爱心":  
            mask = np.array(Image.open(get_resource_path('images/heart.png')))
        
        elif layout == "五角星":
            mask = np.array(Image.open(get_resource_path('images/star.png')))
        
        elif layout == "气泡":
            mask = np.array(Image.open(get_resource_path('images/bubble.png')))
        
        elif layout ==  "长方形":
            mask = np.array(Image.open(get_resource_path('images/rectangle.png')))
        elif layout == "奇异":
            mask = np.array(Image.open(get_resource_path('images/odd.png')))

        elif layout == "四角星":
            mask = np.array(Image.open(get_resource_path('images/four_star.png')))

        elif layout == "圆形":
            mask = np.array(Image.open(get_resource_path('images/circle.png')))

        wc = WordCloud(font_path=get_resource_path('images/Serif.ttf'), background_color='white',  
                       max_words=2000, mask=mask)
        
        filtered_freq_dict = {k: v for k, v in self.freq_dict.items() if v >= self.threshold}  
        
        if not filtered_freq_dict:  
            messagebox.showinfo("信息", "没有足够的词汇满足当前词频阈值")  # 过滤低于阈值的词
            self.status_label.config(text="")
            return
        
        self.freq_dict_draw = dict(sorted(filtered_freq_dict.items(), key=lambda item: item[1], reverse=True))
        wc.generate_from_frequencies(self.freq_dict_draw)  # 生成词云
        self.image = wc.to_image()  
        try:
            self.image.save("./wordcloud.png" )   # 保存词云图片 
        except:
            pass

    # 打开编辑窗口
    def open_edit_window(self):
        """
        打开一个新的窗口来显示和编辑词频。
        """
        self.edit_window = tk.Toplevel(self.root)


        self.edit_window.iconbitmap(default=get_resource_path("images/cloud.ico"))
        width = 1150
        heigh = 650
        screenwidth = self.edit_window.winfo_screenwidth()
        screenheight = self.edit_window.winfo_screenheight()
        self.edit_window.geometry('%dx%d+%d+%d'%(width, heigh, (screenwidth-width)/2, (screenheight-heigh)/2))
        self.edit_window.title("词云生成系统(版权声明@Hsl)")
        
        #添加组件，并展示内容
        self.tree = ttk.Treeview(self.edit_window, columns=("Word", "Frequency"), show="headings")  
        self.tree.heading("Word", text="词语")  
        self.tree.heading("Frequency", text="频率")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)  
        
        # 添加垂直滚动条
        scrollbar = ttk.Scrollbar(self.edit_window, orient="vertical", command=self.tree.yview)  
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)  
        self.tree.configure(yscrollcommand=scrollbar.set)  
        
        # 对词频字典进行排序并更新显示
        self.freq_dict_draw = dict(sorted(self.freq_dict.items(), key=lambda item: item[1], reverse=True))
    
        # 在树形结构中插入排序后的词语和频率
        for word, freq in self.freq_dict_draw.items():  
            self.tree.insert("", tk.END, values=(word, freq))

        # 创建按钮框架
        frame_buttons = tk.Frame(self.edit_window)
        frame_buttons.pack(pady=10)

        # 添加“添加词语”按钮，并绑定事件
        button_add = tk.Button(frame_buttons, text="添加词语", command=self.add_word_dialog, width=15, height=2)
        button_add.pack(side=tk.LEFT, padx=5)

        # 添加“删除词语”按钮，并绑定事件
        button_delete = tk.Button(frame_buttons, text="删除词语", command=self.delete_word, width=15, height=2)
        button_delete.pack(side=tk.LEFT, padx=5)
        
        # 添加“修改词语”按钮，并绑定事件
        button_update_word = tk.Button(frame_buttons, text="修改词语", command=self.update_word_dialog, width=15, height=2)
        button_update_word.pack(side=tk.LEFT, padx=5)

        # 添加“修改词频”按钮，并绑定事件
        button_update_frequency = tk.Button(frame_buttons, text="修改词频", command=self.update_frequency_dialog, width=15, height=2)
        button_update_frequency.pack(side=tk.LEFT, padx=5)

        # 添加“更新词云”按钮，并绑定事件
        button_update = tk.Button(frame_buttons, text="更新词云", command=self.update_from_tree, width=15, height=2)
        button_update.pack(side=tk.LEFT, padx=5)

        # 创建布局和阈值设置的框架
        frame_layout_and_threshold = tk.Frame(self.edit_window)
        frame_layout_and_threshold.pack(pady=5)

        # 创建“词云布局”标签并绑定布局选择框
        label_layout_edit = tk.Label(frame_layout_and_threshold, text="词云布局:")
        label_layout_edit.pack(side=tk.LEFT, padx=5)

        # 词云布局选择框
        self.combo_layout_edit = ttk.Combobox(frame_layout_and_threshold, values=["长方形", "圆形", "爱心", "五角星", "气泡", "奇异", "四角星"])
        self.combo_layout_edit.current(self.combo_layout.current())
        self.combo_layout_edit.pack(side=tk.LEFT, padx=5)

        # 创建“词频阈值”标签
        label_threshold_edit = tk.Label(frame_layout_and_threshold, text="词频阈值:")
        label_threshold_edit.pack(side=tk.LEFT, padx=5)  

        # 词频阈值输入框
        self.entry_threshold_edit = tk.Entry(frame_layout_and_threshold, width=10)
        self.entry_threshold_edit.insert(0, str(self.threshold))
        self.entry_threshold_edit.pack(side=tk.LEFT, padx=5)

        # 创建画布用于显示词云
        self.edit_canvas = tk.Canvas(self.edit_window, width=600, height=500)
        self.edit_canvas.pack(pady=10)

    
    def add_word_dialog(self):
        """
        创建一个模态对话框来添加词语和频率。
        """
        #开启窗口，调整参数
        dialog = tk.Toplevel(self.edit_window)
        dialog.title("添加词语")
        dialog.iconbitmap(default=get_resource_path("images/cloud.ico"))
        width = 250
        heigh = 150
        screenwidth = dialog.winfo_screenwidth()
        screenheight = dialog.winfo_screenheight()
        dialog.geometry('%dx%d+%d+%d'%(width, heigh, (screenwidth-width)/2, (screenheight-heigh)/2))
        dialog.transient(self.edit_window)  
        dialog.grab_set()
        
        #添加词语和频率输入框
        label_word = tk.Label(dialog, text="词语:")
        label_word.grid(row=0, column=0, padx=10, pady=10)
        
        entry_word = tk.Entry(dialog, width=20)
        entry_word.grid(row=0, column=1, padx=10, pady=10)
        
        label_frequency = tk.Label(dialog, text="频率:")
        label_frequency.grid(row=1, column=0, padx=10, pady=10)
        
        entry_frequency = tk.Entry(dialog, width=20)
        entry_frequency.grid(row=1, column=1, padx=10, pady=10)
        
        #添加执行的逻辑
        def on_add():
            word = entry_word.get().strip()
            frequency_str = entry_frequency.get().strip()
            
            if not word or not frequency_str.isdigit():
                messagebox.showerror("错误", "请输入有效的词语和频率")
                return
            
            frequency = int(frequency_str)
            self.tree.insert("", tk.END, values=(word, frequency))
            self.freq_dict[word] = frequency  # 更新词频字典

            # 更新树形结构显示
            for item in self.tree.get_children():
                self.tree.delete(item)

            self.freq_dict_draw = dict(sorted(self.freq_dict.items(), key=lambda item: item[1], reverse=True))

            for word, freq in self.freq_dict_draw.items():  
                self.tree.insert("", tk.END, values=(word, freq))



            dialog.destroy()
        
        button_add = tk.Button(dialog, text="添加", command=on_add)
        button_add.grid(row=2, column=0, columnspan=2, pady=10)
    
    def delete_word(self):
        """
        删除选中的词语及其频率。
        """
        selected_item = self.tree.selection()  
        if selected_item:
            item_values = self.tree.item(selected_item, 'values')  
            word = item_values[0]
            self.tree.delete(selected_item)  
            del self.freq_dict[word]  # 从词频字典中删除对应词语
    
    def update_word_dialog(self):
        """
        创建一个模态对话框来修改词语。
        """
        selected_item = self.tree.selection()  
        if not selected_item:
            messagebox.showwarning("警告", "请选择一个词语进行修改")
            return
        
        item_values = self.tree.item(selected_item, 'values')  
        current_word = item_values[0]
        current_frequency = item_values[1]
        
        #开启窗口调整参数
        dialog = tk.Toplevel(self.edit_window)
        dialog.title("修改词语")
        dialog.iconbitmap(default=get_resource_path("images/cloud.ico"))
        width = 250
        heigh = 100
        screenwidth = dialog.winfo_screenwidth()
        screenheight = dialog.winfo_screenheight()
        dialog.geometry('%dx%d+%d+%d'%(width, heigh, (screenwidth-width)/2, (screenheight-heigh)/2))
        dialog.transient(self.edit_window)  
        dialog.grab_set()
        
        #添加组件
        label_new_word = tk.Label(dialog, text="新词语:")
        label_new_word.grid(row=0, column=0, padx=10, pady=10)
        
        entry_new_word = tk.Entry(dialog, width=20)
        entry_new_word.grid(row=0, column=1, padx=10, pady=10)
        entry_new_word.insert(0, current_word)  # 默认填充当前词语
        
        #删除的逻辑
        def on_update_word():
            new_word = entry_new_word.get().strip()
            
            if not new_word:
                messagebox.showerror("错误", "请输入有效的词语")
                return
            
            self.tree.item(selected_item, values=(new_word, current_frequency))
            self.freq_dict[new_word] = self.freq_dict.pop(current_word)  # 更新词频字典
            
            dialog.destroy()
        
        button_update = tk.Button(dialog, text="修改", command=on_update_word,width=10, height=1)
        button_update.grid(row=1, column=0, columnspan=2, pady=10)
    
    def update_frequency_dialog(self):
        """
        创建一个模态对话框来修改词频。
        """
        selected_item = self.tree.selection()   # 获取树视图中选中的项
        if not selected_item:  # 如果没有选中任何词语，弹出警告框
            messagebox.showwarning("警告", "请选择一个词语进行修改")
            return
        
        # 获取选中项的词语和频率
        item_values = self.tree.item(selected_item, 'values')  
        current_word = item_values[0]
        current_frequency = item_values[1]
        
        # 创建修改词频的对话框
        dialog = tk.Toplevel(self.edit_window)
        dialog.title("修改词频")
        dialog.iconbitmap(default=get_resource_path("images/cloud.ico"))  # 设置图标
        width = 250
        heigh = 100
        screenwidth = dialog.winfo_screenwidth()  # 获取屏幕宽度
        screenheight = dialog.winfo_screenheight()  # 获取屏幕高度
        dialog.geometry('%dx%d+%d+%d' % (width, heigh, (screenwidth - width) / 2, (screenheight - heigh) / 2))  # 设置窗口居中
        dialog.transient(self.edit_window)  # 使得对话框属于主窗口的子窗口
        dialog.grab_set()  # 阻止主窗口与对话框之间的交互
        
        # 添加组件
        label_new_frequency = tk.Label(dialog, text="新频率:")  # 标签：新频率
        label_new_frequency.grid(row=0, column=0, padx=10, pady=10)

        entry_new_frequency = tk.Entry(dialog, width=20)  # 输入框：输入新频率
        entry_new_frequency.grid(row=0, column=1, padx=10, pady=10)
        entry_new_frequency.insert(0, str(current_frequency))  # 默认填充当前频率
        
        #按键逻辑
        def on_update_frequency():  # 获取输入的新频率
            new_frequency_str = entry_new_frequency.get().strip()
            
            # 验证输入的频率是否为有效数字
            if not new_frequency_str.isdigit():
                messagebox.showerror("错误", "请输入有效的频率")
                return
            
            new_frequency = int(new_frequency_str)  # 将输入的频率转换为整数
            
            self.tree.item(selected_item, values=(current_word, new_frequency))
            self.freq_dict[current_word] = new_frequency  # 更新词频字典
            
            # 重新更新树视图中的所有内容，按照频率降序排序
            for item in self.tree.get_children():
                self.tree.delete(item)
                
            self.freq_dict_draw = dict(sorted(self.freq_dict.items(), key=lambda item: item[1], reverse=True))

            for word, freq in self.freq_dict_draw.items():  
                self.tree.insert("", tk.END, values=(word, freq))


            dialog.destroy()  # 关闭对话框
        
        button_update = tk.Button(dialog, text="修改", command=on_update_frequency,width=10, height=1)
        button_update.grid(row=1, column=0, columnspan=2, pady=10)
    
    def update_from_tree(self):
        """
        从树状视图中更新词频字典并重新生成词云。
        """
        self.freq_dict.clear()  # 清空现有的词频字典
        for item in self.tree.get_children():  # 遍历树视图中的所有词语
            values = self.tree.item(item, 'values')  # 获取每个词语和频率
            word, freq = values[0], int(values[1])  # 获取词语和频率
            self.freq_dict[word] = freq  # 更新词频字典

        # 按照频率降序排序词频字典    
        self.freq_dict = dict(sorted(self.freq_dict.items(), key=lambda item: item[1], reverse=True))  
        self.threshold = int(self.entry_threshold_edit.get())  # 获取新的词频阈值
        self.layout = self.combo_layout_edit.get()  # 获取新的词云布局

        self.update_wordcloud_display()  # 更新词云显示
        self.display_wordcloud_in_edit_window()  # 在编辑窗口显示词云图像 
    
    def display_wordcloud_in_edit_window(self):
        """
        在编辑窗口中显示词云图像。
        """
        original_image = self.image  # 获取原始的词云图片

        # 获取原始图片的宽度和高度
        orig_width, orig_height = original_image.size

        # 计算缩放比例（这里以画布的宽度为例，您也可以选择高度）
        max_width = 600
        scale_ratio = max_width / orig_width
        new_height = int(orig_height * scale_ratio)  # 根据比例计算新的高度
 
        # 缩小图片
        resized_image = original_image.resize((int(max_width), new_height), resample=Image.Resampling.LANCZOS)
 
        # 转换为Tkinter可以识别的图片格式
        img_tk = ImageTk.PhotoImage(resized_image)

        if hasattr(self, 'image_on_edit_canvas'):
            self.edit_canvas.delete(self.image_on_edit_canvas)
        
        self.image_on_edit_canvas = self.edit_canvas.create_image(300, 250, image=img_tk)
        self.edit_canvas.image = img_tk
    
    def on_layout_change(self, event):
        """
        当用户更改布局时，重新生成词云并更新显示。
        """
        self.layout = self.combo_layout_edit.get()
        self.update_wordcloud_display()
        self.display_wordcloud_in_edit_window()

    def on_threshold_change(self):
        """
        当用户更改词频阈值时，重新生成词云并更新显示。
        """
        try:
            new_threshold = int(self.entry_threshold_edit.get())
            self.threshold = new_threshold         
            filtered_freq_dict = {k: v for k, v in self.freq_dict.items() if v >= self.threshold}  
            
            if not filtered_freq_dict:  
                messagebox.showinfo("信息", "没有足够的词汇满足当前词频阈值")  
                self.status_label.config(text="")
                return
            #排序
            self.freq_dict = dict(sorted(filtered_freq_dict.items(), key=lambda item: item[1], reverse=True)) 

        except ValueError:
            pass  # 忽略无效输入



if __name__ == "__main__":
    root = tk.Tk()
    app = WordCloudGenerator(root)
    root.mainloop()