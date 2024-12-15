import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QMenuBar, QMenu, QAction, QFileDialog
from PyQt5.QtCore import Qt
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.dates import DateFormatter
import json
import os
from matplotlib.font_manager import FontProperties

# 为Mac系统设置中文字体
if sys.platform.startswith('darwin'):  # Mac系统
    font = FontProperties(fname='fonts/LXGWWenKai-Regular.ttf')  # Mac系统自带的苹方字体
else:  # Windows系统
    font = FontProperties(family='SimHei')

# 设置matplotlib中文字体
plt.rcParams['font.sans-serif'] = ['PingFang HK', 'Arial Unicode MS']  # Mac常用中文字体
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.family'] = 'sans-serif'

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel 数据可视化")
        self.setGeometry(100, 100, 1200, 600)
        
        # 当前打开的文件路径
        self.current_file = None
        
        # 创建菜单栏
        self.create_menu_bar()
        
        # 加载最近打开文件历史
        self.recent_files = self.load_recent_files()
        self.update_recent_files_menu()
        
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 创建水平布局
        layout = QHBoxLayout(main_widget)
        
        # 左侧表格
        self.table = QTableWidget()
        layout.addWidget(self.table)
        
        # 启用多选功能
        self.table.setSelectionMode(QTableWidget.MultiSelection)
        # 改为选择行变化事件
        self.table.itemSelectionChanged.connect(self.on_selection_change)
        
        # 右侧图表
        self.figure = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        
        # 设置布局比例
        layout.setStretch(0, 1)
        layout.setStretch(1, 1)
        
        # 加载数据
        self.load_data()
    
    def create_menu_bar(self):
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu('文件')
        
        # 打开文件动作
        open_action = QAction('打开', self)
        open_action.setShortcut('Ctrl+O')
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)
        
        # 最近打开的文件子菜单
        self.recent_menu = file_menu.addMenu('最近打开')
        
    def load_recent_files(self):
        try:
            config_path = os.path.join(os.path.expanduser('~'), '.excel_viewer_history.json')
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"加载历史记录失败: {str(e)}")
        return []
    
    def save_recent_files(self):
        try:
            config_path = os.path.join(os.path.expanduser('~'), '.excel_viewer_history.json')
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.recent_files, f, ensure_ascii=False)
        except Exception as e:
            print(f"保存历史记录失败: {str(e)}")
    
    def update_recent_files_menu(self):
        self.recent_menu.clear()
        for file_path in self.recent_files:
            action = QAction(os.path.basename(file_path), self)
            action.setStatusTip(file_path)
            action.triggered.connect(lambda checked, path=file_path: self.open_recent_file(path))
            self.recent_menu.addAction(action)
    
    def add_recent_file(self, file_path):
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
        self.recent_files.insert(0, file_path)
        self.recent_files = self.recent_files[:10]  # 只保留最近10个
        self.save_recent_files()
        self.update_recent_files_menu()
    
    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            self.load_excel_file(file_path)
    
    def open_recent_file(self, file_path):
        if os.path.exists(file_path):
            self.load_excel_file(file_path)
        else:
            # 如果文件不存在，从历史记录中移除
            self.recent_files.remove(file_path)
            self.save_recent_files()
            self.update_recent_files_menu()
    
    def load_excel_file(self, file_path):
        try:
            self.current_file = file_path
            df = pd.read_excel(file_path)
            
            # 设置表格
            self.table.setRowCount(len(df))
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels(df.columns)
            
            # 填充表格数据
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    item = QTableWidgetItem(str(df.iloc[i, j]))
                    self.table.setItem(i, j, item)
            
            # 清空图表
            ax = self.figure.gca()
            ax.clear()
            ax.set_xlabel('时间', fontproperties=font)
            ax.set_ylabel('销量', fontproperties=font)
            ax.set_title('点击表格行显示销量趋势', fontproperties=font)
            self.canvas.draw()
            
            # 添加到最近打开文件历史
            self.add_recent_file(file_path)
            
        except Exception as e:
            print(f"加载文件时出错: {str(e)}")
    
    def on_selection_change(self):
        if not self.current_file:
            return
            
        try:
            # 获取所有选中的行
            selected_rows = set(item.row() for item in self.table.selectedItems())
            if not selected_rows:
                return
            
            # 获取数据
            df = pd.read_excel(self.current_file)
            
            # 准备绘图
            ax = self.figure.gca()
            ax.clear()
            
            # 用于存储所有日期范围
            all_dates = []
            
            # 为每个选中的行绘制折线图
            for row in selected_rows:
                row_data = df.iloc[row]
                json_str = row_data['历史数据-junglescout'].replace('&#10;', '').strip()
                history_data = json.loads(json_str)
                
                dates = pd.to_datetime(history_data['days'])
                sales = [0 if x is None else x for x in history_data['sales']]
                
                all_dates.extend(dates)
                
                # 绘制该ASIN的折线图
                ax.plot(dates, sales, '-o', label=f'ASIN: {row_data["ASIN"]}')
            
            # 设置x轴范围为所有数据的最早到最晚日期
            if all_dates:
                min_date = min(all_dates)
                max_date = max(all_dates)
                ax.set_xlim(min_date, max_date)
            
            # 设置图表属性
            ax.set_xlabel('时间', fontproperties=font)
            ax.set_ylabel('销量', fontproperties=font)
            ax.set_title('多产品销量趋势对比', fontproperties=font)
            
            # 设置图例字体
            ax.legend(prop=font)
            
            # 设置x轴时间格式
            ax.xaxis.set_major_formatter(DateFormatter('%Y/%m/%d'))
            plt.xticks(rotation=45)
            
            # 自动调整布局
            self.figure.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            print(f"更新图表时出错: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def load_data(self):
        try:
            # 读取Excel文件
            df = pd.read_excel('示例数据.xlsx')
            
            # 设置表格
            self.table.setRowCount(len(df))
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels(df.columns)
            
            # 填充表格数据
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    item = QTableWidgetItem(str(df.iloc[i, j]))
                    self.table.setItem(i, j, item)
            
            # 初始化时不显示折线图
            ax = self.figure.add_subplot(111)
            ax.clear()
            ax.set_xlabel('时间', fontproperties=font)
            ax.set_ylabel('销量', fontproperties=font)
            ax.set_title('点击表格行显示销量趋势', fontproperties=font)
            self.canvas.draw()
            
        except Exception as e:
            print(f"加载数据时出错: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
