import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import pandas as pd
from datetime import datetime

class AssetManager:
    def __init__(self, root):
        self.root = root
        self.root.title("库房资产管理系统 v1.0")
        
        # 初始化数据库
        self.conn = sqlite3.connect('assets.db')
        self.create_table()
        
        # 创建界面组件
        self.create_widgets()
        self.load_data()

    def create_table(self):
        """创建数据库表结构"""
        cursor = self.conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS assets
                          (id INTEGER PRIMARY KEY AUTOINCREMENT,
                           asset_name TEXT NOT NULL,
                           asset_type TEXT,
                           quantity INTEGER DEFAULT 0,
                           location TEXT,
                           last_update TIMESTAMP)''')
        self.conn.commit()

    def create_widgets(self):
        """创建图形界面组件"""
        # 工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="入库", command=self.add_inbound).grid(row=0, column=0, padx=2)
        ttk.Button(toolbar, text="出库", command=self.add_outbound).grid(row=0, column=1, padx=2)
        ttk.Button(toolbar, text="导入Excel", command=self.import_excel).grid(row=0, column=2, padx=2)
        ttk.Button(toolbar, text="导出Excel", command=self.export_excel).grid(row=0, column=3, padx=2)
        ttk.Button(toolbar, text="生成模板", command=self.generate_template).grid(row=0, column=4, padx=2)

        # 数据表格
        self.tree = ttk.Treeview(self.root, columns=('ID', '资产名称', '类型', '数量', '位置', '最后更新'), show='headings')
        self.tree.heading('ID', text='ID')
        self.tree.heading('资产名称', text='资产名称')
        self.tree.heading('类型', text='类型')
        self.tree.heading('数量', text='数量')
        self.tree.heading('位置', text='位置')
        self.tree.heading('最后更新', text='最后更新')
        self.tree.pack(fill=tk.BOTH, expand=True)

        # 状态栏
        self.status = ttk.Label(self.root, text="就绪", relief=tk.SUNKEN)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    def load_data(self):
        """加载数据到表格"""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM assets")
        for row in cursor.fetchall():
            self.tree.insert('', tk.END, values=row)

    # ----------------- 核心功能方法 -----------------
    def add_inbound(self):
        """入库操作"""
        self._show_edit_dialog("入库", is_outbound=False)

    def add_outbound(self):
        """出库操作"""
        self._show_edit_dialog("出库", is_outbound=True)

    def _show_edit_dialog(self, title, is_outbound=False):
        """显示入库/出库对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        
        ttk.Label(dialog, text="资产名称:").grid(row=0, column=0, padx=5, pady=5)
        asset_name = ttk.Entry(dialog)
        asset_name.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="数量:").grid(row=1, column=0, padx=5, pady=5)
        quantity = ttk.Entry(dialog)
        quantity.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="类型:").grid(row=2, column=0, padx=5, pady=5)
        asset_type = ttk.Combobox(dialog, values=["设备", "耗材", "工具", "其他"])
        asset_type.grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="位置:").grid(row=3, column=0, padx=5, pady=5)
        location = ttk.Entry(dialog)
        location.grid(row=3, column=1, padx=5, pady=5)

        def save():
            try:
                qty = int(quantity.get())
                if is_outbound:
                    qty = -abs(qty)
                else:
                    qty = abs(qty)
                
                cursor = self.conn.cursor()
                cursor.execute('''INSERT INTO assets 
                                  (asset_name, asset_type, quantity, location, last_update)
                                  VALUES (?, ?, ?, ?, ?)''',
                                  (asset_name.get(), asset_type.get(), 
                                   qty, location.get(), datetime.now()))
                self.conn.commit()
                self.load_data()
                dialog.destroy()
                self.status.config(text=f"{title}操作成功")
            except Exception as e:
                messagebox.showerror("错误", f"操作失败: {str(e)}")

        ttk.Button(dialog, text="保存", command=save).grid(row=4, columnspan=2, pady=10)

    def import_excel(self):
        """从Excel导入数据"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                cursor = self.conn.cursor()
                
                for _, row in df.iterrows():
                    cursor.execute('''INSERT INTO assets 
                                    (asset_name, asset_type, quantity, location, last_update)
                                    VALUES (?, ?, ?, ?, ?)''',
                                    (row['资产名称'], row['类型'], 
                                     row['数量'], row['位置'], datetime.now()))
                self.conn.commit()
                self.load_data()
                self.status.config(text=f"成功导入 {len(df)} 条记录")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}")

    def export_excel(self):
        """导出数据到Excel"""
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            try:
                df = pd.read_sql("SELECT * FROM assets", self.conn)
                df.to_excel(file_path, index=False)
                self.status.config(text=f"数据已导出到 {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")

    def generate_template(self):
        """生成Excel模板"""
        template = pd.DataFrame(columns=["资产名称", "类型", "数量", "位置"])
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel模板", "*.xlsx")])
        if file_path:
            try:
                template.to_excel(file_path, index=False)
                self.status.config(text=f"模板已生成到 {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"生成失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AssetManager(root)
    root.mainloop()