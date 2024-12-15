import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QMenuBar, QMenu, QAction, QFileDialog, QLabel
from PyQt5.QtCore import Qt
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.dates import DateFormatter
import json
import os
from matplotlib.font_manager import FontProperties
import requests
from PyQt5.QtGui import QPixmap
import hashlib
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar

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
        
        # 加载用户设置
        self.settings = self.load_settings()
        
        # 创建菜单栏
        self.create_menu_bar()
        
        # 加载最近打开文件历史
        self.recent_files = self.load_recent_files()
        self.update_recent_files_menu()
        
        # 创建主窗口部件和UI
        self.setup_ui()
        
        # 初始化颜色设置
        self.init_colors()
        
        # 根据设置决定是否自动加载上次的文件，如果不加载则显示空白界面
        if self.settings.get('auto_load_last_file', False) and self.recent_files:
            last_file = self.recent_files[0]
            if os.path.exists(last_file):
                self.load_excel_file(last_file)
        else:
            self.load_data()  # 显示空白界面
    
    def setup_ui(self):
        """设置UI组件"""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 创建主布局
        layout = QHBoxLayout(main_widget)
        
        # 左侧表格
        self.table = QTableWidget()
        layout.addWidget(self.table)
        
        # 启用多选功能和行为设置
        self.table.setSelectionMode(QTableWidget.MultiSelection)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.itemSelectionChanged.connect(self.on_selection_change)
        
        # 右侧布局（包含图表和工具栏）
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # 创建图表
        self.figure = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.figure)
        
        # 添加工具栏
        self.toolbar = NavigationToolbar(self.canvas, right_widget)
        
        # 将工具栏和画布添加到右侧布局
        right_layout.addWidget(self.toolbar)
        right_layout.addWidget(self.canvas)
        
        # 将右侧部件添加到主布局
        layout.addWidget(right_widget)
        
        # 设置布局比例
        layout.setStretch(0, 1)
        layout.setStretch(1, 1)
        
        # 创建imgs文件夹（如果不存在）
        if not os.path.exists('imgs'):
            os.makedirs('imgs')
        
        # 设置表格的行高
        self.table.verticalHeader().setDefaultSectionSize(100)
    
    def init_colors(self):
        """初始化颜色设置"""
        self.colors = [
            '#1f77b4',  # 蓝色
            '#ff7f0e',  # 橙色
            '#2ca02c',  # 绿色
            '#d62728',  # 红色
            '#9467bd',  # 紫色
            '#8c564b',  # 棕色
            '#e377c2',  # 粉色
            '#7f7f7f',  # 灰色
            '#bcbd22',  # 黄绿色
            '#17becf',  # 青色
            '#ff9896',  # 浅红色
            '#98df8a',  # 浅绿色
            '#c5b0d5',  # 浅紫色
            '#c49c94',  # 浅棕色
            '#f7b6d2',  # 浅粉色
            '#dbdb8d',  # 浅黄色
            '#9edae5',  # 浅青色
            '#ad494a',  # 深红色
            '#8c6d31',  # 深黄色
            '#bd9e39'   # 金色
        ]
        self.asin_colors = {}
        self.used_color_indices = set()
    
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
        
        # 设置菜单
        settings_menu = menubar.addMenu('设置')
        
        # 自动加载上次文件的选项
        self.auto_load_action = QAction('自动加载上次文件', self)
        self.auto_load_action.setCheckable(True)
        self.auto_load_action.setChecked(self.settings.get('auto_load_last_file', False))
        self.auto_load_action.triggered.connect(self.toggle_auto_load)
        settings_menu.addAction(self.auto_load_action)
        
        # 添加开始时间设置选项
        self.start_time_action = QAction('开始时间为上架时间', self)
        self.start_time_action.setCheckable(True)
        self.start_time_action.setChecked(self.settings.get('start_from_launch_date', True))
        self.start_time_action.triggered.connect(self.toggle_start_time)
        settings_menu.addAction(self.start_time_action)
    
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
            # 如果文件不存在，从史记录中移除
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
            
            # 获取图片链接列的索引
            image_col_index = df.columns.get_loc("图片链接")
            
            # 填充表格数据
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    item = QTableWidgetItem(str(df.iloc[i, j]))
                    self.table.setItem(i, j, item)
                    
                    # 如果是图片链接列，检查是否已有缓存���片
                    if j == image_col_index:
                        image_url = df.iloc[i, j]
                        # 使用URL的MD5作为文件名
                        filename = hashlib.md5(image_url.encode()).hexdigest() + '.jpg'
                        local_path = os.path.join('imgs', filename)
                        
                        # 如果图片已存在，直接显示
                        if os.path.exists(local_path):
                            image_label = self.create_image_label(local_path)
                            self.table.setCellWidget(i, j, image_label)
            
            # 添加表格项目点击事件（仅用于处理未下载的图片）
            self.table.itemClicked.connect(self.on_item_clicked)
            
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
            import traceback
            traceback.print_exc()
    
    def get_next_color(self):
        """获取下一个未使用的颜色"""
        # 如果所有颜色都用完了，重置使用记录
        if len(self.used_color_indices) >= len(self.colors):
            self.used_color_indices.clear()
            
        # 找到第一个未使用的颜色
        for i in range(len(self.colors)):
            if i not in self.used_color_indices:
                self.used_color_indices.add(i)
                return self.colors[i]
                
        # 如果没有找到（理论上不会发生），返回第一个颜色
        return self.colors[0]
    
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
            launch_dates = []  # 存储所有产品的上架时间
            first_data_dates = []  # 存储所有产品历史数据的第一个时间
            max_dates = []  # 存储所有产品的最新数据时间
            
            # 为每个选中的行绘制折线图
            for row in selected_rows:
                row_data = df.iloc[row]
                asin = row_data["ASIN"]
                
                if asin not in self.asin_colors:
                    self.asin_colors[asin] = self.get_next_color()
                
                json_str = row_data['历史数据-junglescout'].replace('&#10;', '').strip()
                history_data = json.loads(json_str)
                
                # 获取上架时间
                launch_date = pd.to_datetime(row_data['上架日期'])
                launch_dates.append(launch_date)
                
                # 转换历史数据日期
                dates = pd.to_datetime(history_data['days'], format='%Y/%m/%d')
                first_data_dates.append(dates.min())  # 记录历史数据的第一个时间
                sales = [0 if x is None else x for x in history_data['sales']]
                
                # 获取最新数据时间
                latest_date = dates.max()
                max_dates.append(latest_date)
                
                # 绘制折线图
                ax.plot(dates, sales, '-o', label=f'ASIN: {asin}', 
                       color=self.asin_colors[asin])
            
            # 根据设置选择开始时间
            if self.settings.get('start_from_launch_date', True):
                start_date = min(launch_dates)
            else:
                start_date = min(first_data_dates)
            
            # 设置x轴范围
            if max_dates:
                ax.set_xlim(start_date, max(max_dates))
            
            # 设置图表属性
            ax.set_xlabel('时间', fontproperties=font)
            ax.set_ylabel('销量', fontproperties=font)
            ax.set_title('多产品销量趋势对比', fontproperties=font)
            
            # 设置图例字体
            ax.legend(prop=font)
            
            # 设置x轴时间格式
            ax.xaxis.set_major_formatter(DateFormatter('%Y/%m/%d'))
            ax.xaxis.set_tick_params(labelrotation=-45)  # 设置 x 轴标签旋转 -45 度
            
            # 自动调整布局
            self.figure.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            print(f"更新图表时出错: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def load_data(self):
        """移除默认加载数据的行为"""
        # 初始化空表格
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        
        # 初始化空图表
        ax = self.figure.add_subplot(111)
        ax.clear()
        ax.set_xlabel('时间', fontproperties=font)
        ax.set_ylabel('销量', fontproperties=font)
        ax.set_title('点击表格行显示销量趋势', fontproperties=font)
        self.canvas.draw()
    
    def download_image(self, url):
        """下载图片并返回本地路径"""
        try:
            # 使用URL的MD5作为文件名
            filename = hashlib.md5(url.encode()).hexdigest() + '.jpg'
            local_path = os.path.join('imgs', filename)
            
            # 如果图片已存在，直接返回路径
            if os.path.exists(local_path):
                return local_path
            
            # 下载图片
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                with open(local_path, 'wb') as f:
                    f.write(response.content)
                return local_path
        except Exception as e:
            print(f"下载图片失败: {str(e)}")
        return None

    def create_image_label(self, image_path):
        """创建图片标签并设置图片"""
        label = QLabel()
        pixmap = QPixmap(image_path)
        scaled_pixmap = pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        label.setPixmap(scaled_pixmap)
        return label

    def on_item_clicked(self, item):
        try:
            row = item.row()
            col = item.column()
            df = pd.read_excel(self.current_file)
            
            # 只处理图片链接列的点击
            if col == df.columns.get_loc("图片链接"):
                # 如果单元格中已经有图片，不需要处理
                if self.table.cellWidget(row, col) is not None:
                    return
                    
                # 获取图片链接并下载
                image_url = df.iloc[row, col]
                image_path = self.download_image(image_url)
                if image_path:
                    image_label = self.create_image_label(image_path)
                    self.table.setCellWidget(row, col, image_label)
                
        except Exception as e:
            print(f"加载图片时出错: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def load_settings(self):
        """加载用户设置"""
        try:
            settings_path = os.path.join(os.path.expanduser('~'), '.excel_viewer_settings.json')
            if os.path.exists(settings_path):
                with open(settings_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"加载设置失败: {str(e)}")
        return {}
    
    def save_settings(self):
        """保存用户设置"""
        try:
            settings_path = os.path.join(os.path.expanduser('~'), '.excel_viewer_settings.json')
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False)
        except Exception as e:
            print(f"保存设置失败: {str(e)}")
    
    def toggle_auto_load(self, checked):
        """切换自动加载设置"""
        self.settings['auto_load_last_file'] = checked
        self.save_settings()
    
    def toggle_start_time(self, checked):
        """切换开始时间设置"""
        self.settings['start_from_launch_date'] = checked
        self.save_settings()
        # 如果有选中的行，立即更新图表
        if self.table.selectedItems():
            self.on_selection_change()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
