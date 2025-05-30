import sys

from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
                             QCheckBox, QLabel, QComboBox, QTabWidget, QDialog, QMessageBox, QSplitter, QHeaderView,
                             QGridLayout, QSpinBox, QDoubleSpinBox, QLineEdit, QGroupBox, QTextEdit, QScrollArea,
                             QMenu, QAction, QFileDialog)
from PyQt5.QtCore import Qt
import pandas as pd
import akshare as ak
import mplfinance as mpf
from datetime import datetime, timedelta
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        def parse_percentage(s):
            s = s.strip().replace("%", "")
            try:
                return float(s) / 100
            except ValueError:
                return float('-inf')

        try:
            if "%" in self.text() and "%" in other.text():
                return parse_percentage(self.text()) < parse_percentage(other.text())
            return float(self.text()) < float(other.text())
        except ValueError:
            return super().__lt__(other)


class ChartWindow(QDialog):
    def __init__(self, stock_code):
        super().__init__()
        self.stock_code = stock_code
        self.initUI()

    def initUI(self):
        # 设置matplotlib中文字体
        plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # macOS
        # 如果是 Windows 系统，使用：
        # plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.rcParams['axes.unicode_minus'] = False
        self.setWindowTitle(f'股票图表 - {self.stock_code}')
        self.setGeometry(300, 300, 800, 600)

        layout = QVBoxLayout()

        # 图表类型选择
        self.chart_type = QComboBox()
        self.chart_type.addItems(['K线图', '分时图', 'MACD', 'KDJ', 'RSI'])
        self.chart_type.currentTextChanged.connect(self.update_chart)

        layout.addWidget(self.chart_type)

        # 图表显示区域
        self.figure = Figure(figsize=(8, 6))
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)

        self.setLayout(layout)
        self.update_chart(self.chart_type.currentText())

    def update_chart(self, chart_type):
        self.figure.clear()
        if chart_type == 'K线图':
            self.plot_candlestick()
        elif chart_type == '分时图':
            self.plot_timeline()
        elif chart_type == 'MACD':
            self.plot_macd()
        elif chart_type == 'KDJ':
            self.plot_kdj()
        elif chart_type == 'RSI':
            self.plot_rsi()
        self.canvas.draw()

    def plot_timeline(self):
        try:
            # 清除当前图形
            self.figure.clear()

            # 获取分时数据
            df = ak.stock_zh_a_hist_min_em(
                symbol=self.stock_code,
                period='1',
                start_date=(datetime.now()).strftime('%Y%m%d'),
                end_date=datetime.now().strftime('%Y%m%d')
            )

            # 检查数据是否为空
            if df.empty:
                print("没有获取到数据")
                return

            # 打印数据，用于调试
            print("收盘价范围:", df['收盘'].min(), "-", df['收盘'].max())
            print("成交量范围:", df['成交量'].min(), "-", df['成交量'].max())

            # 准备数据
            df['时间'] = pd.to_datetime(df['时间'])
            df.set_index('时间', inplace=True)

            # 检查处理后的数据
            print("数据点数量:", len(df))

            # 创建子图
            ax1 = self.figure.add_subplot(111)

            # 绘制分时价格并返回线条对象
            line1 = ax1.plot(df.index, df['收盘'], 'b-', label='Price')
            print("是否创建了价格线:", len(line1) > 0)

            # 创建成交量子图
            ax2 = ax1.twinx()
            bars = ax2.bar(df.index, df['成交量'], alpha=0.3, color='g', label='Volume')
            print("是否创建了成交量柱:", len(bars) > 0)

            # 设置坐标轴范围
            ax1.set_xlim(df.index.min(), df.index.max())
            y_margin = (df['收盘'].max() - df['收盘'].min()) * 0.1
            ax1.set_ylim(df['收盘'].min() - y_margin, df['收盘'].max() + y_margin)

            # 设置标题和标签
            ax1.set_title(f'{self.stock_code} Timeline')
            ax1.set_xlabel('Time')
            ax1.set_ylabel('Price', color='b')
            ax2.set_ylabel('Volume', color='g')

            # 设置x轴时间格式
            ax1.tick_params(axis='x', rotation=45)
            ax1.xaxis.set_major_locator(mdates.MinuteLocator(interval=30))
            ax1.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))

            # 添加网格
            ax1.grid(True, linestyle='--', alpha=0.6)

            # 自动调整布局
            self.figure.tight_layout()

            # 添加图例
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right')

            # 刷新画布
            self.canvas.draw()

        except Exception as e:
            print(f"Failed to plot timeline: {e}")
            import traceback
            traceback.print_exc()

    def plot_macd(self):
        try:
            # 获取数据
            df = ak.stock_zh_a_hist(
                symbol=self.stock_code,
                period="daily",
                start_date=(datetime.now() - timedelta(days=120)).strftime('%Y%m%d'),
                end_date=datetime.now().strftime('%Y%m%d')
            )

            # 计算MACD
            close = df['收盘'].astype(float)
            exp1 = close.ewm(span=12, adjust=False).mean()
            exp2 = close.ewm(span=26, adjust=False).mean()
            macd = exp1 - exp2
            signal = macd.ewm(span=9, adjust=False).mean()
            histogram = macd - signal

            # 创建子图
            gs = self.figure.add_gridspec(2, 1, height_ratios=[2, 1])
            ax1 = self.figure.add_subplot(gs[0])
            ax2 = self.figure.add_subplot(gs[1])

            # 绘制价格
            ax1.plot(df['日期'], close, label='收盘价')

            # 绘制MACD
            dates = df['日期']
            ax2.plot(dates, macd, label='MACD')
            ax2.plot(dates, signal, label='Signal')

            # 绘制直方图
            for i in range(len(dates)):
                if histogram[i] >= 0:
                    ax2.bar(dates[i], histogram[i], color='r', width=0.8)
                else:
                    ax2.bar(dates[i], histogram[i], color='g', width=0.8)

            # 设置标题和图例
            ax1.set_title(f'{self.stock_code} MACD分析')
            ax1.legend()
            ax2.legend()

            # 设置x轴格式
            ax1.tick_params(axis='x', rotation=45)
            ax2.tick_params(axis='x', rotation=45)

        except Exception as e:
            print(f"绘制MACD失败: {e}")
            import traceback
            traceback.print_exc()

    def plot_kdj(self):
        try:
            # 获取数据
            df = ak.stock_zh_a_hist(
                symbol=self.stock_code,
                period="daily",
                start_date=(datetime.now() - timedelta(days=120)).strftime('%Y%m%d'),
                end_date=datetime.now().strftime('%Y%m%d')
            )

            # 计算KDJ
            low_list = df['最低'].rolling(window=9, min_periods=9).min()
            high_list = df['最高'].rolling(window=9, min_periods=9).max()
            rsv = (df['收盘'] - low_list) / (high_list - low_list) * 100

            k = rsv.ewm(com=2, adjust=False).mean()
            d = k.ewm(com=2, adjust=False).mean()
            j = 3 * k - 2 * d

            # 创建子图
            gs = self.figure.add_gridspec(2, 1, height_ratios=[2, 1])
            ax1 = self.figure.add_subplot(gs[0])
            ax2 = self.figure.add_subplot(gs[1])

            # 绘制价格
            ax1.plot(df['日期'], df['收盘'], label='收盘价')

            # 绘制KDJ
            ax2.plot(df['日期'], k, label='K')
            ax2.plot(df['日期'], d, label='D')
            ax2.plot(df['日期'], j, label='J')

            # 设置标题和图例
            ax1.set_title(f'{self.stock_code} KDJ分析')
            ax1.legend()
            ax2.legend()

            # 设置x轴格式
            ax1.tick_params(axis='x', rotation=45)
            ax2.tick_params(axis='x', rotation=45)

        except Exception as e:
            print(f"绘制KDJ失败: {e}")
            import traceback
            traceback.print_exc()

    def plot_rsi(self):
        try:
            # 获取数据
            df = ak.stock_zh_a_hist(
                symbol=self.stock_code,
                period="daily",
                start_date=(datetime.now() - timedelta(days=120)).strftime('%Y%m%d'),
                end_date=datetime.now().strftime('%Y%m%d')
            )

            # 计算RSI
            def calculate_rsi(data, periods=14):
                delta = data.diff()
                gain = (delta.where(delta > 0, 0)).rolling(window=periods).mean()
                loss = (-delta.where(delta < 0, 0)).rolling(window=periods).mean()
                rs = gain / loss
                return 100 - (100 / (1 + rs))

            rsi_6 = calculate_rsi(df['收盘'].astype(float), 6)
            rsi_12 = calculate_rsi(df['收盘'].astype(float), 12)
            rsi_24 = calculate_rsi(df['收盘'].astype(float), 24)

            # 创建子图
            gs = self.figure.add_gridspec(2, 1, height_ratios=[2, 1])
            ax1 = self.figure.add_subplot(gs[0])
            ax2 = self.figure.add_subplot(gs[1])

            # 绘制价格
            ax1.plot(df['日期'], df['收盘'], label='收盘价')

            # 绘制RSI
            ax2.plot(df['日期'], rsi_6, label='RSI6')
            ax2.plot(df['日期'], rsi_12, label='RSI12')
            ax2.plot(df['日期'], rsi_24, label='RSI24')

            # 添加RSI的参考线
            ax2.axhline(y=80, color='r', linestyle='--', alpha=0.3)
            ax2.axhline(y=20, color='g', linestyle='--', alpha=0.3)
            ax2.axhline(y=50, color='b', linestyle='--', alpha=0.3)

            # 设置Y轴范围
            ax2.set_ylim([0, 100])

            # 设置标题和图例
            ax1.set_title(f'{self.stock_code} RSI分析')
            ax1.legend()
            ax2.legend()

            # 设置x轴格式
            ax1.tick_params(axis='x', rotation=45)
            ax2.tick_params(axis='x', rotation=45)

        except Exception as e:
            print(f"绘制RSI失败: {e}")
            import traceback
            traceback.print_exc()

    def plot_candlestick(self):
        try:
            stock_data = ak.stock_zh_a_hist(
                symbol=self.stock_code,
                period="daily",
                start_date=(datetime.now() - timedelta(days=60)).strftime('%Y%m%d'),
                end_date=datetime.now().strftime('%Y%m%d')
            )

            # 准备数据
            stock_data.index = pd.to_datetime(stock_data['日期'])
            stock_data = stock_data.rename(columns={
                '开盘': 'Open', '最高': 'High',
                '最低': 'Low', '收盘': 'Close',
                '成交量': 'Volume'
            })

            # 创建两个子图，一个用于K线，一个用于成交量
            self.figure.clear()

            # 创建网格布局，调整成交量图的高度比例
            gs = self.figure.add_gridspec(2, 1, height_ratios=[3, 1])

            # K线图子图
            ax1 = self.figure.add_subplot(gs[0])
            # 成交量图子图
            ax2 = self.figure.add_subplot(gs[1])

            # 设置K线图样式
            mc = mpf.make_marketcolors(up='red', down='green',
                                       edge='inherit',
                                       wick='inherit',
                                       volume='in')
            s = mpf.make_mpf_style(marketcolors=mc)

            # 绘制K线图和成交量
            mpf.plot(stock_data,
                     type='candle',
                     style=s,
                     ax=ax1,
                     volume=ax2,
                     tight_layout=True)

            # 设置标题
            self.figure.suptitle(f'{self.stock_code} K线图')

        except Exception as e:
            print(f"绘制K线图失败: {e}")
            import traceback
            traceback.print_exc()

class StockScreener(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ma_trend_cache = {}  # 添加缓存字典
        self.market_trend_cache = {}  # 添加大盘趋势缓存
        self.initUI()

    def initUI(self):
        self.setWindowTitle('股票筛选器')
        self.setGeometry(100, 100, 1200, 800)

        # 主布局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()

        # 顶部控制区
        control_layout = QHBoxLayout()

        # 刷新按钮
        refresh_btn = QPushButton('刷新数据')
        refresh_btn.clicked.connect(self.refresh_data)
        control_layout.addWidget(refresh_btn)

        # 筛选按钮
        filter_btn = QPushButton('执行筛选')
        filter_btn.clicked.connect(self.apply_filter)
        control_layout.addWidget(filter_btn)

        # 清空状态按钮
        clear_btn = QPushButton('清空条件')
        clear_btn.clicked.connect(self.clear_filter_conditions)
        control_layout.addWidget(clear_btn)

        # 设置默认值按钮
        default_btn = QPushButton('默认条件')
        default_btn.clicked.connect(self.set_default_conditions)
        control_layout.addWidget(default_btn)

        # 添加涨停票分析按钮
        analyze_limit_up_btn = QPushButton('涨停票分析')
        analyze_limit_up_btn.clicked.connect(self.show_limit_up_analysis)
        control_layout.addWidget(analyze_limit_up_btn)

        # 添加大盘分析按钮
        analyze_market_btn = QPushButton('大盘分析')
        analyze_market_btn.clicked.connect(self.show_market_analysis)
        control_layout.addWidget(analyze_market_btn)

        # 添加资金流向分析按钮
        money_flow_btn = QPushButton('资金流向分析')
        money_flow_btn.clicked.connect(self.show_money_flow_analysis)
        control_layout.addWidget(money_flow_btn)

        # 添加资金净流入排序按钮
        money_flow_rank_btn = QPushButton('资金净流入排序')
        money_flow_rank_btn.clicked.connect(self.show_money_flow_rank)
        control_layout.addWidget(money_flow_rank_btn)

        # 添加主力排名按钮
        main_fund_rank_btn = QPushButton('主力排名')
        main_fund_rank_btn.clicked.connect(self.show_main_fund_rank)
        control_layout.addWidget(main_fund_rank_btn)

        layout.addLayout(control_layout)

        # 添加两个增长股票显示区域
        growing_stocks_group = QGroupBox("增长股票列表")
        growing_stocks_layout = QVBoxLayout()

        # 所有增长股票显示区域
        all_growing_layout = QHBoxLayout()
        all_growing_layout.addWidget(QLabel('所有增长股票:'))
        self.growing_stocks_edit = QLineEdit()
        self.growing_stocks_edit.setReadOnly(True)
        all_growing_layout.addWidget(self.growing_stocks_edit)

        # 主板增长股票显示区域
        main_board_layout = QHBoxLayout()
        main_board_layout.addWidget(QLabel('主板增长股票:'))
        self.main_board_stocks_edit = QLineEdit()
        self.main_board_stocks_edit.setReadOnly(True)
        main_board_layout.addWidget(self.main_board_stocks_edit)

        # 添加复制按钮
        copy_buttons_layout = QHBoxLayout()
        copy_all_btn = QPushButton('复制所有增长股票')
        copy_all_btn.clicked.connect(lambda: self.copy_stocks_text(self.growing_stocks_edit.text()))
        copy_main_btn = QPushButton('复制主板增长股票')
        copy_main_btn.clicked.connect(lambda: self.copy_stocks_text(self.main_board_stocks_edit.text()))

        copy_buttons_layout.addWidget(copy_all_btn)
        copy_buttons_layout.addWidget(copy_main_btn)

        growing_stocks_layout.addLayout(all_growing_layout)
        growing_stocks_layout.addLayout(main_board_layout)
        growing_stocks_layout.addLayout(copy_buttons_layout)

        # 添加新按钮到control_layout
        show_ma_btn = QPushButton('显示主板多头向上股票')
        show_ma_btn.clicked.connect(self.show_ma_stocks)
        control_layout.addWidget(show_ma_btn)

        show_ma_up_btn = QPushButton('显示主板多头向上且上涨股票')
        show_ma_up_btn.clicked.connect(self.show_ma_up_stocks)
        control_layout.addWidget(show_ma_up_btn)

        # 添加新的显示区域
        ma_stocks_layout = QHBoxLayout()
        ma_stocks_layout.addWidget(QLabel('多头向上股票:'))
        self.ma_stocks_edit = QLineEdit()
        self.ma_stocks_edit.setReadOnly(True)
        ma_stocks_layout.addWidget(self.ma_stocks_edit)

        ma_up_stocks_layout = QHBoxLayout()
        ma_up_stocks_layout.addWidget(QLabel('多头向上且上涨股票:'))
        self.ma_up_stocks_edit = QLineEdit()
        self.ma_up_stocks_edit.setReadOnly(True)
        ma_up_stocks_layout.addWidget(self.ma_up_stocks_edit)

        growing_stocks_layout.addLayout(ma_stocks_layout)
        growing_stocks_layout.addLayout(ma_up_stocks_layout)

        growing_stocks_group.setLayout(growing_stocks_layout)
        layout.addWidget(growing_stocks_group)

        # 筛选条件区域
        self.filter_conditions = []
        self.setup_filter_conditions(layout)

        # 创建上下分割的布局
        splitter = QSplitter(Qt.Vertical)

        # 原有的股票列表
        self.stock_table = QTableWidget()
        self.stock_table.setColumnCount(6)
        self.stock_table.setHorizontalHeaderLabels(['代码', '名称', '现价', '涨跌幅', '换手率', '量比'])
        self.stock_table.cellClicked.connect(self.show_stock_charts)
        splitter.addWidget(self.stock_table)

        # 添加筛选结果表格
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(8)
        self.result_table.setHorizontalHeaderLabels([
            '代码', '名称', '现价', '涨跌幅',
            '换手率', '量比', '成交量', '成交额'
        ])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.result_table.horizontalHeader().setStretchLastSection(True)
        self.result_table.verticalHeader().setVisible(False)
        # 为结果表格也添加点击事件
        self.result_table.cellClicked.connect(self.show_stock_charts)
        splitter.addWidget(self.result_table)

        # 将分割器添加到主布局
        layout.addWidget(splitter)

        main_widget.setLayout(layout)

        # 初始加载数据
        self.refresh_data()

        # 为表格添加右键菜单
        self.stock_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.stock_table.customContextMenuRequested.connect(self.show_context_menu)
        self.result_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.result_table.customContextMenuRequested.connect(self.show_context_menu)

    def copy_stocks_text(self, text):
        """复制股票列表到剪贴板"""
        if text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            QMessageBox.information(self, '提示', '已复制到剪贴板')
        else:
            QMessageBox.warning(self, '提示', '没有可复制的内容')

    def clear_filter_conditions(self):
        # 清空所有筛选条件
        self.top_n_spin.setValue(0)
        self.turnover_min.setValue(0)
        self.turnover_max.setValue(0)
        self.price_change_min.setValue(-20)
        self.price_change_max.setValue(20)
        self.volume_ratio_min.setValue(0)
        self.volume_ratio_max.setValue(0)
        self.price_min.setValue(0)
        self.price_max.setValue(10000)
        self.hot_stocks_n.setValue(0)
        self.market_cap_min.setValue(0)
        self.market_cap_max.setValue(10000)
        self.months_spin.setValue(0)  # 设为0
        self.limit_up_times.setValue(0)  # 设为0

        # 清空复选框
        self.volume_increase_cb.setChecked(False)
        self.remove_green_cb.setChecked(False)
        self.remove_limit_up_cb.setChecked(False)
        self.ma_alignment_cb.setChecked(False)
        self.macd_golden_cb.setChecked(False)
        self.kdj_golden_cb.setChecked(False)

    def set_default_conditions(self):
        # 设置默认筛选条件
        self.top_n_spin.setValue(0)
        self.turnover_min.setValue(3)
        self.turnover_max.setValue(10)
        self.price_change_min.setValue(3)
        self.price_change_max.setValue(10)
        self.volume_ratio_min.setValue(1)
        self.volume_ratio_max.setValue(10)
        self.price_min.setValue(1)
        self.price_max.setValue(5)
        self.hot_stocks_n.setValue(0)
        self.market_cap_min.setValue(50)
        self.market_cap_max.setValue(300)
        self.months_spin.setValue(1)
        self.limit_up_times.setValue(1)

        # 设置默认复选框状态
        self.volume_increase_cb.setChecked(False)
        self.remove_green_cb.setChecked(True)
        self.remove_limit_up_cb.setChecked(False)
        self.ma_alignment_cb.setChecked(False)
        self.macd_golden_cb.setChecked(False)
        self.kdj_golden_cb.setChecked(False)

    def setup_filter_conditions(self, layout):
        # Add search controls at the beginning
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('输入股票代码或名称进行搜索(多个用逗号分隔)')
        search_btn = QPushButton('搜索')
        search_btn.clicked.connect(self.search_stocks)

        search_layout.addWidget(self.search_input)
        search_layout.addWidget(search_btn)
        layout.addLayout(search_layout)
        # 创建筛选条件区域
        filter_widget = QWidget()
        filter_layout = QGridLayout()
        current_row = 0

        # 1. 排行榜前N只股票
        filter_layout.addWidget(QLabel('排行榜前N只:'), current_row, 0)
        self.top_n_spin = QSpinBox()
        self.top_n_spin.setRange(0, 5000)
        self.top_n_spin.setValue(0)
        filter_layout.addWidget(self.top_n_spin, current_row, 1)

        # 2. 换手率范围
        filter_layout.addWidget(QLabel('换手率范围(%):'), current_row, 2)
        self.turnover_min = QDoubleSpinBox()
        self.turnover_max = QDoubleSpinBox()
        self.turnover_min.setRange(0, 100)
        self.turnover_max.setRange(0, 100)
        self.turnover_max.setValue(10)
        filter_layout.addWidget(self.turnover_min, current_row, 3)
        filter_layout.addWidget(QLabel('至'), current_row, 4)
        filter_layout.addWidget(self.turnover_max, current_row, 5)

        current_row += 1

        # 3. 涨幅范围
        filter_layout.addWidget(QLabel('涨幅范围(%):'), current_row, 0)
        self.price_change_min = QDoubleSpinBox()
        self.price_change_max = QDoubleSpinBox()
        self.price_change_min.setRange(-20, 20)
        self.price_change_max.setRange(-20, 20)
        self.price_change_min.setValue(3)
        self.price_change_max.setValue(10)
        filter_layout.addWidget(self.price_change_min, current_row, 1)
        filter_layout.addWidget(QLabel('至'), current_row, 2)
        filter_layout.addWidget(self.price_change_max, current_row, 3)

        # 4. 量比范围
        filter_layout.addWidget(QLabel('量比范围:'), current_row, 4)
        self.volume_ratio_min = QDoubleSpinBox()
        self.volume_ratio_max = QDoubleSpinBox()
        self.volume_ratio_min.setRange(0, 100)
        self.volume_ratio_max.setRange(0, 100)
        self.volume_ratio_min.setValue(1)
        filter_layout.addWidget(self.volume_ratio_min, current_row, 5)

        current_row += 1

        # 5-7. 复选框条件
        self.volume_increase_cb = QCheckBox('成交量持续放大')
        self.remove_green_cb = QCheckBox('删除绿盘')
        self.remove_limit_up_cb = QCheckBox('删除涨停')
        filter_layout.addWidget(self.volume_increase_cb, current_row, 0)
        filter_layout.addWidget(self.remove_green_cb, current_row, 1)
        filter_layout.addWidget(self.remove_limit_up_cb, current_row, 2)

        current_row += 1

        # 8. 股价范围
        filter_layout.addWidget(QLabel('股价范围:'), current_row, 0)
        self.price_min = QDoubleSpinBox()
        self.price_max = QDoubleSpinBox()
        self.price_min.setRange(0, 10000)
        self.price_max.setRange(0, 10000)
        self.price_min.setValue(1)
        self.price_max.setValue(5)
        filter_layout.addWidget(self.price_min, current_row, 1)
        filter_layout.addWidget(QLabel('至'), current_row, 2)
        filter_layout.addWidget(self.price_max, current_row, 3)

        # 9. 热门个股
        filter_layout.addWidget(QLabel('热门个股前N:'), current_row, 4)
        self.hot_stocks_n = QSpinBox()
        self.hot_stocks_n.setRange(0, 5000)
        self.hot_stocks_n.setValue(0)
        filter_layout.addWidget(self.hot_stocks_n, current_row, 5)

        current_row += 1

        # 10. 市值范围
        filter_layout.addWidget(QLabel('市值范围(亿):'), current_row, 0)
        self.market_cap_min = QDoubleSpinBox()
        self.market_cap_max = QDoubleSpinBox()
        self.market_cap_min.setRange(0, 10000)
        self.market_cap_max.setRange(0, 10000)
        self.market_cap_min.setValue(50)
        self.market_cap_max.setValue(300)
        filter_layout.addWidget(self.market_cap_min, current_row, 1)
        filter_layout.addWidget(QLabel('至'), current_row, 2)
        filter_layout.addWidget(self.market_cap_max, current_row, 3)

        current_row += 1

        # 11. 涨停次数
        filter_layout.addWidget(QLabel('最近涨停:'), current_row, 0)
        self.months_spin = QSpinBox()
        self.months_spin.setRange(0, 12)  # 改为从0开始
        self.months_spin.setValue(0)
        filter_layout.addWidget(self.months_spin, current_row, 1)
        filter_layout.addWidget(QLabel('个月内第'), current_row, 2)
        self.limit_up_times = QSpinBox()
        self.limit_up_times.setRange(0, 10)  # 改为从0开始
        self.limit_up_times.setValue(0)
        filter_layout.addWidget(self.limit_up_times, current_row, 3)
        filter_layout.addWidget(QLabel('次涨停'), current_row, 4)

        current_row += 1

        # 12-13. 技术指标
        self.ma_alignment_cb = QCheckBox('均线多头排列')
        self.macd_golden_cb = QCheckBox('MACD金叉')
        self.kdj_golden_cb = QCheckBox('KDJ金叉')
        filter_layout.addWidget(self.ma_alignment_cb, current_row, 0)
        filter_layout.addWidget(self.macd_golden_cb, current_row, 1)
        filter_layout.addWidget(self.kdj_golden_cb, current_row, 2)

        filter_widget.setLayout(filter_layout)
        layout.addWidget(filter_widget)

    def search_stocks(self):
        search_text = self.search_input.text().strip()
        if not search_text:
            return

        try:
            df = ak.stock_zh_a_spot_em()

            # 将中英文逗号统一替换为英文逗号后再分割
            search_items = [item.strip() for item in search_text.replace('，', ',').split(',')]

            all_matches = pd.DataFrame()

            for search_item in search_items:
                if not search_item:
                    continue

                # 精确匹配
                exact_matches = df[
                    (df['代码'] == search_item) |
                    (df['名称'] == search_item)
                    ]

                # 如果没有精确匹配，进行模糊搜索
                if exact_matches.empty:
                    fuzzy_matches = df[
                        (df['代码'].str.contains(search_item, case=False, na=False)) |
                        (df['名称'].str.contains(search_item, case=False, na=False))
                        ]
                    if not fuzzy_matches.empty:
                        all_matches = pd.concat([all_matches, fuzzy_matches])
                else:
                    all_matches = pd.concat([all_matches, exact_matches])

            if not all_matches.empty:
                # 去重
                all_matches = all_matches.drop_duplicates(subset=['代码'])
                self.show_filtered_results(all_matches)
                QMessageBox.information(self, "搜索结果", f"找到 {len(all_matches)} 个相关结果")
            else:
                QMessageBox.information(self, "搜索结果", "未找到匹配的股票")

        except Exception as e:
            print(f"搜索股票失败: {e}")
            QMessageBox.warning(self, "错误", "搜索过程中发生错误")

    def refresh_data(self):
        self.clear_ma_trend_cache()
        try:
            df = ak.stock_zh_a_spot_em()
            self.stock_table.setRowCount(len(df))

            for i, (_, row) in enumerate(df.iterrows()):
                self.stock_table.setItem(i, 0, QTableWidgetItem(str(row['代码'])))
                self.stock_table.setItem(i, 1, QTableWidgetItem(str(row['名称'])))
                self.stock_table.setItem(i, 2, NumericTableWidgetItem(str(row['最新价'])))
                self.stock_table.setItem(i, 3, NumericTableWidgetItem(str(row['涨跌幅'])))
                self.stock_table.setItem(i, 4, NumericTableWidgetItem(str(row['换手率'])))
                self.stock_table.setItem(i, 5, NumericTableWidgetItem(str(row['量比'])))

            # 筛选并显示增长股票
            growing_stocks = df[df['涨跌幅'] > 0].sort_values(by='涨跌幅', ascending=False)
            growing_stocks_text = ', '.join([f"{row['名称']}"
                                                 for _, row in growing_stocks.iterrows()])
            # 筛选主板增长股票（排除创业板、科创板和北交所）
            main_board_stocks = growing_stocks[
                (~growing_stocks['代码'].str.startswith('300')) &  # 排除创业板
                (~growing_stocks['代码'].str.startswith('688')) &  # 排除科创板
                (~growing_stocks['代码'].str.startswith('430')) &  # 排除北交所
                (~growing_stocks['代码'].str.startswith('689')) &  # 排除科创板
                (~growing_stocks['代码'].str.startswith('830')) &  # 排除北交所
                (~growing_stocks['代码'].str.startswith('839'))  # 排除北交所
                ]

            # 更新所有增长股票显示
            growing_stocks_text = ', '.join([
                f"{row['名称']}"
                for _, row in growing_stocks.iterrows()
            ])
            self.growing_stocks_edit.setText(growing_stocks_text)

            # 更新主板增长股票显示
            main_board_stocks_text = ', '.join([
                f"{row['名称']}"
                for _, row in main_board_stocks.iterrows()
            ])
            self.main_board_stocks_edit.setText(main_board_stocks_text)
            # self.growing_stocks_edit.setText(growing_stocks_text)
            # 筛选主板涨停票
            limit_up_stocks = df[
                (df['涨跌幅'] >= 9.5) &  # 涨停幅度为10%
                (~df['代码'].str.startswith('300')) &  # 排除创业板
                (~df['代码'].str.startswith('688')) &  # 排除科创板
                (~df['代码'].str.startswith('430')) &  # 排除北交所
                (~df['代码'].str.startswith('689')) &  # 排除科创板
                (~df['代码'].str.startswith('830')) &  # 排除北交所
                (~df['代码'].str.startswith('839'))  # 排除北交所
                ]
            limit_up_stocks_text = ', '.join([f"{row['名称']}" for _, row in limit_up_stocks.iterrows()])


            print(f"所有增长股票显示: {growing_stocks_text}\n")
            print(f"主板增长股票: {main_board_stocks_text}\n")
            print(f"主板涨停股票: {limit_up_stocks_text}\n")

            # # 筛选主板股票
            # all_main_board_stocks = df[
            #     (~df['代码'].str.startswith('300')) &
            #     (~df['代码'].str.startswith('688')) &
            #     (~df['代码'].str.startswith('430')) &
            #     (~df['代码'].str.startswith('689')) &
            #     (~df['代码'].str.startswith('830')) &
            #     (~df['代码'].str.startswith('839'))
            #     ]

        except Exception as e:
            print(f"刷新数据失败: {e}")

    def check_ma_trend(self, stock_code):
        # 检查缓存
        if stock_code in self.ma_trend_cache:
            return self.ma_trend_cache[stock_code]

        try:
            # 获取60天的历史数据
            hist_data = ak.stock_zh_a_hist(
                symbol=stock_code,
                period="daily",
                start_date=(datetime.now() - timedelta(days=60)).strftime('%Y%m%d'),
                end_date=datetime.now().strftime('%Y%m%d')
            )

            if hist_data.empty:
                return (False, None)

            # 计算均线
            ma5 = hist_data['收盘'].rolling(5).mean()
            ma10 = hist_data['收盘'].rolling(10).mean()
            ma20 = hist_data['收盘'].rolling(20).mean()
            ma30 = hist_data['收盘'].rolling(30).mean()

            # 计算均线的变化率（斜率）
            ma5_slope = ma5.diff()
            ma10_slope = ma10.diff()
            ma20_slope = ma20.diff()
            ma30_slope = ma30.diff()

            # 获取最近几天的数据来判断拐点
            lookback_days = 3  # 向前看3天的趋势变化

            # 判断均线拐头向上的条件：
            # 1. 最近一天的斜率为正（向上）
            # 2. 之前的斜率为负（向下）或接近于0
            # 3. 斜率由负变正，表示出现拐点
            is_turning_up = (
                # MA5拐头向上
                (ma5_slope.iloc[-1] > 0) and  # 最近一天向上
                (ma5_slope.iloc[-lookback_days:-1].mean() <= 0) and  # 之前趋势向下

                # MA10拐头向上
                (ma10_slope.iloc[-1] > 0) and
                (ma10_slope.iloc[-lookback_days:-1].mean() <= 0) and

                # MA20拐头向上
                (ma20_slope.iloc[-1] > 0) and
                (ma20_slope.iloc[-lookback_days:-1].mean() <= 0)
            )

            # 判断多头排列
            is_bullish = (
                # 当前多头排列
                (ma5.iloc[-1] > ma10.iloc[-1] > ma20.iloc[-1] > ma30.iloc[-1]) and
                # 均线斜率都为正
                (ma5_slope.iloc[-1] > 0) and
                (ma10_slope.iloc[-1] > 0) and
                (ma20_slope.iloc[-1] > 0) and
                (ma30_slope.iloc[-1] > 0)
            )

            # 获取最新数据
            latest = hist_data.iloc[-1]

            result = (is_turning_up or is_bullish, latest)
            self.ma_trend_cache[stock_code] = result
            return result

        except Exception as e:
            print(f"处理股票{stock_code}时出错: {e}")
            return (False, None)

    def check_vol_price_up(self, stock_code, hist_data, days=3):
        """
        检查连续多天量价齐升

        Parameters:
            stock_code: 股票代码
            hist_data: 历史数据
            days: 需要比较的天数（从今天往前推），默认为2

        Returns:
            bool: 是否连续量价齐升
        """
        try:
            if len(hist_data) < days + 1:  # 确保有足够的数据
                return False

            # 逐日比较
            for i in range(days):
                today = hist_data.iloc[-(i + 1)]  # 当前日
                yesterday = hist_data.iloc[-(i + 2)]  # 前一日

                # 判断量价齐升
                price_up = today['收盘'] > yesterday['收盘']
                volume_up = today['成交量'] > yesterday['成交量']

                # 如果任一天不满足量价齐升，返回False
                if not (price_up and volume_up):
                    return False

            return True

        except Exception as e:
            print(f"检查量价齐升时出错 {stock_code}: {e}")
            return False

    def show_ma_stocks(self):
        try:
            df = ak.stock_zh_a_spot_em()
            # 筛选主板股票
            main_board_stocks = df[
                (~df['代码'].str.startswith('300')) &
                (~df['代码'].str.startswith('688')) &
                (~df['代码'].str.startswith('430')) &
                (~df['代码'].str.startswith('689')) &
                (~df['代码'].str.startswith('830')) &
                (~df['代码'].str.startswith('839'))
                ]

            # 创建六个列表分别存储不同类型的股票
            ma_up_not_limit = []  # 多头向上且上涨但非涨停
            ma_up_not_limit_vol = []  # 多头向上且上涨但非涨停且量价齐升
            ma_up_limit = []  # 多头向上且涨停
            ma_up_limit_vol = []  # 多头向上且涨停且量价齐升
            ma_down = []  # 多头向上且下跌
            ma_down_vol = []  # 多头向上且下跌且量价齐升

            for _, stock in main_board_stocks.iterrows():
                try:
                    is_bullish, latest = self.check_ma_trend(stock['代码'])
                    if is_bullish:
                        # 获取历史数据用于量价齐升判断
                        hist_data = ak.stock_zh_a_hist(
                            symbol=stock['代码'],
                            period="daily",
                            start_date=(datetime.now() - timedelta(days=5)).strftime('%Y%m%d'),
                            end_date=datetime.now().strftime('%Y%m%d')
                        )

                        is_vol_price_up = self.check_vol_price_up(stock['代码'], hist_data)

                        # 涨跌幅判断
                        change_pct = stock['涨跌幅']
                        if change_pct > 0:
                            if change_pct >= 9.5:  # 涨停判断
                                ma_up_limit.append(f"{stock['名称']}")
                                if is_vol_price_up:
                                    ma_up_limit_vol.append(f"{stock['名称']}")
                            else:
                                ma_up_not_limit.append(f"{stock['名称']}")
                                if is_vol_price_up:
                                    ma_up_not_limit_vol.append(f"{stock['名称']}")
                        else:
                            ma_down.append(f"{stock['名称']}")
                            if is_vol_price_up:
                                ma_down_vol.append(f"{stock['名称']}")
                except Exception as e:
                    print(f"处理股票 {stock['代码']} {stock['名称']} 时出错: {e}")
                    continue

            # 生成显示文本
            ma_up_not_limit_text = ', '.join(ma_up_not_limit)
            ma_up_not_limit_vol_text = ', '.join(ma_up_not_limit_vol)
            ma_up_limit_text = ', '.join(ma_up_limit)
            ma_up_limit_vol_text = ', '.join(ma_up_limit_vol)
            ma_down_text = ', '.join(ma_down)
            ma_down_vol_text = ', '.join(ma_down_vol)

            # 合并所有股票用于显示在文本框中
            all_ma_stocks = ma_up_not_limit + ma_up_limit + ma_down
            all_ma_stocks_text = ', '.join(all_ma_stocks)
            self.ma_stocks_edit.setText(all_ma_stocks_text)

            # 打印详细分类信息
            print("\n=== 主板多头向上股票分类 ===")
            print(f"多头向上且上涨未涨停({len(ma_up_not_limit)}只): {ma_up_not_limit_text}")
            print(f"其中量价齐升({len(ma_up_not_limit_vol)}只): {ma_up_not_limit_vol_text}")

            print(f"\n多头向上且涨停({len(ma_up_limit)}只): {ma_up_limit_text}")
            print(f"其中量价齐升({len(ma_up_limit_vol)}只): {ma_up_limit_vol_text}")

            print(f"\n多头向上且下跌({len(ma_down)}只): {ma_down_text}")
            print(f"其中量价齐升({len(ma_down_vol)}只): {ma_down_vol_text}")
            print("=" * 50)

        except Exception as e:
            print(f"获取多头向上股票失败: {e}")

    def show_ma_up_stocks(self):
        try:
            df = ak.stock_zh_a_spot_em()
            # 筛选主板上涨股票
            main_board_up_stocks = df[
                (df['涨跌幅'] > 0) &
                (~df['代码'].str.startswith('300')) &
                (~df['代码'].str.startswith('688')) &
                (~df['代码'].str.startswith('430')) &
                (~df['代码'].str.startswith('689')) &
                (~df['代码'].str.startswith('830')) &
                (~df['代码'].str.startswith('839'))
                ]

            ma_up_stocks = []
            for _, stock in main_board_up_stocks.iterrows():
                try:
                    is_bullish, latest = self.check_ma_trend(stock['代码'])
                    if is_bullish:
                        ma_up_stocks.append(f"{stock['名称']}")
                except Exception as e:
                    print(f"处理股票 {stock['代码']} {stock['名称']} 时出错: {e}")
                    continue

            # 更新显示
            ma_up_stocks_text = ', '.join(ma_up_stocks)
            self.ma_up_stocks_edit.setText(ma_up_stocks_text)
            print(f"\n主板多头向上且上涨股票: {ma_up_stocks_text}")

        except Exception as e:
            print(f"获取多头向上且上涨股票失败: {e}")

    # 可以添加一个清除缓存的方法
    def clear_ma_trend_cache(self):
        self.ma_trend_cache.clear()

    def apply_filter(self):
        # 进行筛选
        filtered_stocks = self.filter_stocks()
        if filtered_stocks is not None:
            self.show_filtered_results(filtered_stocks)

    def filter_stocks(self):
        try:
            # 获取实时行情数据
            df = ak.stock_zh_a_spot_em()

            # 应用筛选条件
            # 换手率范围
            df = df[(df['换手率'] >= self.turnover_min.value()) &
                    (df['换手率'] <= self.turnover_max.value())]

            # 涨幅范围
            df = df[(df['涨跌幅'] >= self.price_change_min.value()) &
                    (df['涨跌幅'] <= self.price_change_max.value())]

            # 量比范围
            df = df[df['量比'] >= self.volume_ratio_min.value()]
            if self.volume_ratio_max.value() > 0:
                df = df[df['量比'] <= self.volume_ratio_max.value()]

            # 股价范围
            df = df[(df['最新价'] >= self.price_min.value()) &
                    (df['最新价'] <= self.price_max.value())]

            # 市值范围（需要将亿元转换为元）
            min_cap = self.market_cap_min.value() * 100000000
            max_cap = self.market_cap_max.value() * 100000000
            df = df[(df['总市值'] >= min_cap) & (df['总市值'] <= max_cap)]

            # 复选框条件
            if self.remove_green_cb.isChecked():
                df = df[df['涨跌幅'] > 0]

            if self.remove_limit_up_cb.isChecked():
                df = df[df['涨跌幅'] < 9.5]

            if self.volume_increase_cb.isChecked():
                volume_increasing_stocks = []
                for index, row in df.iterrows():
                    stock_code = row['代码']
                    hist_data = ak.stock_zh_a_hist_min_em(
                        symbol=stock_code,
                        period='1',
                        start_date=(datetime.now()).strftime('%Y%m%d'),
                        end_date=datetime.now().strftime('%Y%m%d')
                    )
                    if not hist_data.empty:
                        recent_volumes = hist_data['成交量'].tail(3).values
                        if len(recent_volumes) == 3 and all(
                                recent_volumes[i] > recent_volumes[i - 1] for i in range(1, 3)):
                            volume_increasing_stocks.append(stock_code)
                df = df[df['代码'].isin(volume_increasing_stocks)]

            # 技术指标
            if self.ma_alignment_cb.isChecked() or self.macd_golden_cb.isChecked() or self.kdj_golden_cb.isChecked():
                technical_stocks = []
                for index, row in df.iterrows():
                    stock_code = row['代码']
                    hist_data = ak.stock_zh_a_hist(symbol=stock_code, period="daily",
                                                   start_date=(datetime.now() - timedelta(days=60)).strftime('%Y%m%d'),
                                                   end_date=datetime.now().strftime('%Y%m%d'))

                    if not hist_data.empty:
                        # 计算技术指标
                        if self.ma_alignment_cb.isChecked():
                            # 计算MA5, MA10, MA20
                            ma5 = hist_data['收盘'].rolling(5).mean().iloc[-1]
                            ma10 = hist_data['收盘'].rolling(10).mean().iloc[-1]
                            ma20 = hist_data['收盘'].rolling(20).mean().iloc[-1]
                            if not (ma5 > ma10 > ma20):
                                continue

                        if self.macd_golden_cb.isChecked():
                            # 计算MACD
                            exp1 = hist_data['收盘'].ewm(span=12, adjust=False).mean()
                            exp2 = hist_data['收盘'].ewm(span=26, adjust=False).mean()
                            macd = exp1 - exp2
                            signal = macd.ewm(span=9, adjust=False).mean()
                            if not (macd.iloc[-1] > signal.iloc[-1] and macd.iloc[-2] <= signal.iloc[-2]):
                                continue

                        if self.kdj_golden_cb.isChecked():
                            # 计算KDJ
                            low_list = hist_data['最低'].rolling(9, min_periods=9).min()
                            high_list = hist_data['最高'].rolling(9, min_periods=9).max()
                            rsv = (hist_data['收盘'] - low_list) / (high_list - low_list) * 100
                            k = rsv.ewm(com=2).mean()
                            d = k.ewm(com=2).mean()
                            if not (k.iloc[-1] > d.iloc[-1] and k.iloc[-2] <= d.iloc[-2]):
                                continue

                    technical_stocks.append(stock_code)

                df = df[df['代码'].isin(technical_stocks)]

            # 获取热门个股
            if self.hot_stocks_n.value() > 0:
                df = df.nlargest(self.hot_stocks_n.value(), '成交额')

            # 行业龙头筛选
            if self.top_n_spin.value() > 0:  # 只有当值大于0时才筛选
                industry_leaders = []
                for industry in df['所属行业'].unique():
                    industry_stocks = df[df['所属行业'] == industry]
                    leaders = industry_stocks.nlargest(self.top_n_spin.value(), '市值')
                    industry_leaders.extend(leaders['代码'].tolist())
                df = df[df['代码'].isin(industry_leaders)]

            # 涨停次数分析
            if self.limit_up_times.value() > 0 and self.months_spin.value() > 0:  # 只有当两个值都大于0时才进行筛选
                limit_up_stocks = []
                months = self.months_spin.value()
                required_times = self.limit_up_times.value()

                start_date = (datetime.now() - timedelta(days=30 * months)).strftime('%Y%m%d')
                end_date = datetime.now().strftime('%Y%m%d')

                for index, row in df.iterrows():
                    stock_code = row['代码']
                    hist_data = ak.stock_zh_a_hist(symbol=stock_code, period="daily",
                                                   start_date=start_date,
                                                   end_date=end_date)

                    if not hist_data.empty:
                        # 计算涨停次数
                        limit_up_count = len(hist_data[hist_data['涨跌幅'] >= 9.5])
                        if limit_up_count == required_times:
                            limit_up_stocks.append(stock_code)

                df = df[df['代码'].isin(limit_up_stocks)]

            return df

        except Exception as e:
            print(f"筛选股票失败: {e}")
            import traceback
            traceback.print_exc()
            return None

    def analyze_trading_signals(self, df):
        """分析股票交易信号"""
        try:
            def get_price_position(code, current_price):
                try:
                    # 获取历史数据
                    hist_data = ak.stock_zh_a_hist(
                        symbol=code,
                        period="daily",
                        start_date=(datetime.now() - timedelta(days=120)).strftime('%Y%m%d'),
                        end_date=datetime.now().strftime('%Y%m%d')
                    )
                    
                    if hist_data.empty:
                        return "未知"
                    
                    # 计算相对位置
                    price_max = hist_data['收盘'].max()
                    price_min = hist_data['收盘'].min()
                    price_range = price_max - price_min
                    
                    if price_range == 0:
                        return "未知"
                    
                    position = (current_price - price_min) / price_range * 100
                    
                    if position < 30:
                        return "低位"
                    elif position > 70:
                        return "高位"
                    else:
                        return "中位"
                    
                except Exception as e:
                    print(f"获取价格位置失败: {e}")
                    return "未知"

            results = []
            
            for _, row in df.iterrows():
                try:
                    # 初始化评分变量
                    rating_score = 0
                    risk_score = 0  # 确保在这里初始化
                    
                    price = row['最新价']
                    price_change = row['涨跌幅']
                    volume_ratio = row['量比']
                    turnover = row['换手率']
                    position = get_price_position(row['代码'], price)
                    
                    # 分析建议
                    regular_advice = ""
                    xu_advice = ""
                    regular_reason = ""
                    xu_reason = ""
                    
                    # 根据技术指标和市场表现计算评分
                    if price_change > 5:
                        rating_score += 2
                    elif price_change > 2:
                        rating_score += 1
                        
                    if volume_ratio > 2:
                        rating_score += 1
                        
                    if turnover > 5:
                        rating_score += 1
                        
                    if position == "低位":
                        rating_score += 2
                    elif position == "高位":
                        risk_score += 2
                    
                    # 生成建议
                    if rating_score >= 4 and risk_score <= 1:
                        final_advice = "建议买入"
                    elif rating_score >= 2 and risk_score <= 2:
                        final_advice = "可以关注"
                    elif risk_score >= 3:
                        final_advice = "注意风险"
                    else:
                        final_advice = "建议观望"
                    
                    results.append({
                        '代码': row['代码'],
                        '名称': row['名称'],
                        '现价': price,
                        '涨跌幅': f"{price_change:.2f}%",
                        '量比': f"{volume_ratio:.2f}",
                        '换手率': f"{turnover:.2f}%",
                        '位置': position,
                        '建议': final_advice,
                        '评分': rating_score,
                        '风险分': risk_score
                    })
                    
                except Exception as e:
                    print(f"分析股票 {row['代码']} 失败: {e}")
                    continue
            
            return pd.DataFrame(results)
            
        except Exception as e:
            print(f"分析交易信号失败: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()

    def show_filtered_results(self, df):
        # 获取分析结果
        analysis_df = self.analyze_trading_signals(df)

        # 清空现有表格内容
        self.result_table.clearContents()
        self.result_table.setRowCount(0)

        # 更新列定义
        columns = ['代码', '名称', '现价', '涨跌幅', '量比', '换手率', '位置',
                   '建议', '评分', '风险分']
        self.result_table.setColumnCount(len(columns))
        self.result_table.setHorizontalHeaderLabels(columns)
        
        # 启用排序
        self.result_table.setSortingEnabled(True)

        # 设置表格样式
        self.result_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #E5E5E5;
                background-color: white;
                alternate-background-color: #F5F5F5;
            }
            QHeaderView::section {
                background-color: #F0F0F0;
                padding: 4px;
                border: 1px solid #E0E0E0;
                font-weight: bold;
            }
        """)
        self.result_table.setAlternatingRowColors(True)

        # 临时禁用排序以提高性能
        self.result_table.setSortingEnabled(False)

        # 填充分析结果
        for index, row in analysis_df.iterrows():
            current_row = self.result_table.rowCount()
            self.result_table.insertRow(current_row)

            for col_idx, column in enumerate(columns):
                # 根据列类型创建不同的表格项
                if column in ['现价', '涨跌幅', '量比', '换手率', '评分', '风险分']:
                    item = NumericTableWidgetItem(str(row[column]))
                else:
                    item = QTableWidgetItem(str(row[column]))
                
                item.setTextAlignment(Qt.AlignCenter)

                # 设置涨跌幅颜色
                if column == '涨跌幅':
                    try:
                        change = float(str(row[column]).replace('%', ''))
                        if change > 0:
                            item.setForeground(QBrush(QColor('#FF4444')))
                        elif change < 0:
                            item.setForeground(QBrush(QColor('#00AA00')))
                    except:
                        pass

                # 设置建议颜色
                if column == '建议':
                    if '买入' in str(row[column]) or '加仓' in str(row[column]):
                        item.setForeground(QBrush(QColor('#FF4444')))
                    elif '卖出' in str(row[column]) or '出局' in str(row[column]):
                        item.setForeground(QBrush(QColor('#00AA00')))
                    elif '观望' in str(row[column]) or '持有' in str(row[column]):
                        item.setForeground(QBrush(QColor('#0088FF')))

                self.result_table.setItem(current_row, col_idx, item)

            # 设置行背景色
            if '买入' in row['建议'] or '加仓' in row['建议']:
                for col in range(len(columns)):
                    self.result_table.item(current_row, col).setBackground(QBrush(QColor('#FFEEEE')))
            elif '卖出' in row['建议'] or '出局' in row['建议']:
                for col in range(len(columns)):
                    self.result_table.item(current_row, col).setBackground(QBrush(QColor('#EEFFEE')))

        # 重新启用排序
        self.result_table.setSortingEnabled(True)

        # 调整列宽度
        self.result_table.resizeColumnsToContents()

        # 设置最小列宽
        min_width = 60
        for i in range(self.result_table.columnCount()):
            if self.result_table.columnWidth(i) < min_width:
                self.result_table.setColumnWidth(i, min_width)

    def show_stock_charts(self, row, col):
        stock_code = self.stock_table.item(row, 0).text()
        chart_window = ChartWindow(stock_code)
        chart_window.exec_()

    def analyze_limit_up_stocks(self):
        """分析涨停票并预测趋势"""
        try:
            # 获取实时行情数据
            df = ak.stock_zh_a_spot_em()
            
            # 获取上证指数数据
            sh_index = ak.stock_zh_index_daily_em(symbol="sh000001")
            
            # 获取行业信息
            industry_df = ak.stock_board_industry_name_em()
            
            # 筛选涨停股票
            limit_up_stocks = df[
                (df['涨跌幅'] >= 9.5) &  # 涨停
                (~df['代码'].str.startswith('300')) &  # 排除创业板
                (~df['代码'].str.startswith('688')) &  # 排除科创板
                (~df['代码'].str.startswith('430')) &  # 排除北交所
                (~df['代码'].str.startswith('689')) &  # 排除科创板
                (~df['代码'].str.startswith('830')) &  # 排除北交所
                (~df['代码'].str.startswith('839'))  # 排除北交所
            ]
            
            # 获取行业资金流向
            industry_flow = ak.stock_sector_fund_flow_rank()
            
            # 获取热点消息
            news = ak.stock_news_em()
            
            # 分析结果列表
            analysis_results = []
            
            for _, stock in limit_up_stocks.iterrows():
                try:
                    # 获取历史数据
                    hist_data = ak.stock_zh_a_hist(
                        symbol=stock['代码'],
                        period="daily",
                        start_date=(datetime.now() - timedelta(days=30)).strftime('%Y%m%d'),
                        end_date=datetime.now().strftime('%Y%m%d')
                    )
                    
                    if hist_data.empty:
                        continue
                    
                    # 确保列名正确
                    required_columns = ['收盘', '开盘', '最高', '最低', '成交量']
                    for col in required_columns:
                        if col not in hist_data.columns:
                            print(f"股票 {stock['代码']} 缺少必要的列: {col}")
                            continue
                    
                    # 计算技术指标
                    ma5 = hist_data['收盘'].rolling(5).mean()
                    ma10 = hist_data['收盘'].rolling(10).mean()
                    ma20 = hist_data['收盘'].rolling(20).mean()
                    
                    # 计算RSI
                    delta = hist_data['收盘'].diff()
                    gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
                    loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
                    rs = gain / loss
                    rsi = 100 - (100 / (1 + rs))
                    
                    # 计算MACD
                    exp1 = hist_data['收盘'].ewm(span=12, adjust=False).mean()
                    exp2 = hist_data['收盘'].ewm(span=26, adjust=False).mean()
                    macd = exp1 - exp2
                    signal = macd.ewm(span=9, adjust=False).mean()
                    
                    # 获取最新数据
                    latest = hist_data.iloc[-1]
                    
                    # 分析特征
                    features = []
                    
                    # 1. 连续涨停次数
                    limit_up_count = 0
                    for i in range(len(hist_data)-1, -1, -1):
                        if hist_data.iloc[i]['涨跌幅'] >= 9.5:
                            limit_up_count += 1
                        else:
                            break
                    features.append(f"连续涨停{limit_up_count}次")
                    
                    # 2. 量能特征
                    if stock['量比'] > 2:
                        features.append("放量涨停")
                    elif stock['量比'] < 0.8:
                        features.append("缩量涨停")
                    
                    # 3. 均线特征
                    if ma5.iloc[-1] > ma10.iloc[-1] > ma20.iloc[-1]:
                        features.append("均线多头排列")
                    elif ma5.iloc[-1] < ma10.iloc[-1] < ma20.iloc[-1]:
                        features.append("均线空头排列")
                    
                    # 4. RSI特征
                    if rsi.iloc[-1] > 80:
                        features.append("RSI超买")
                    elif rsi.iloc[-1] < 20:
                        features.append("RSI超卖")
                    
                    # 5. MACD特征
                    if macd.iloc[-1] > signal.iloc[-1] and macd.iloc[-2] <= signal.iloc[-2]:
                        features.append("MACD金叉")
                    elif macd.iloc[-1] < signal.iloc[-1] and macd.iloc[-2] >= signal.iloc[-2]:
                        features.append("MACD死叉")
                    
                    # 6. 行业资金流向
                    try:
                        # 获取股票所属行业
                        stock_industry = industry_df[industry_df['代码'] == stock['代码']]['板块名称'].iloc[0]
                        industry_flow_data = industry_flow[industry_flow['行业'] == stock_industry]
                        if not industry_flow_data.empty:
                            flow_value = industry_flow_data.iloc[0]['主力净流入-净额']
                            if flow_value > 0:
                                features.append(f"{stock_industry}资金净流入")
                            else:
                                features.append(f"{stock_industry}资金净流出")
                    except:
                        pass
                    
                    # 7. 相关消息
                    stock_news = news[news['代码'] == stock['代码']] if '代码' in news.columns else pd.DataFrame()
                    if not stock_news.empty:
                        features.append(f"相关消息{len(stock_news)}条")
                    
                    # 预测趋势
                    trend_prediction = self.predict_trend(
                        hist_data, ma5, ma10, ma20, rsi, macd, signal
                    )
                    
                    # 分析原因
                    reasons = []
                    
                    # 1. 连续涨停分析
                    if limit_up_count >= 2:
                        reasons.append(f"连续{limit_up_count}个涨停，说明上涨动能强")
                    
                    # 2. 量能分析
                    if stock['量比'] > 2:
                        reasons.append("放量涨停说明资金关注度高")
                    elif stock['量比'] < 0.8:
                        reasons.append("缩量涨停需要谨慎对待")
                    
                    # 3. 均线系统分析
                    if ma5.iloc[-1] > ma10.iloc[-1] > ma20.iloc[-1]:
                        ma5_slope = (ma5.iloc[-1] - ma5.iloc[-5]) / ma5.iloc[-5] * 100
                        if ma5_slope > 2:
                            reasons.append("均线多头排列且MA5快速上扬，趋势强劲")
                        else:
                            reasons.append("均线多头排列，趋势向好")
                    elif ma5.iloc[-1] < ma10.iloc[-1] < ma20.iloc[-1]:
                        reasons.append("均线空头排列，需要观察突破情况")
                    
                    # 4. RSI分析
                    if rsi.iloc[-1] > 80:
                        reasons.append("RSI超买，注意回调风险")
                    elif rsi.iloc[-1] < 30:
                        reasons.append("RSI超卖，可能存在反弹机会")
                    elif 50 < rsi.iloc[-1] < 70:
                        reasons.append("RSI处于强势区间，走势健康")
                    
                    # 5. MACD分析
                    if macd.iloc[-1] > signal.iloc[-1] and macd.iloc[-2] <= signal.iloc[-2]:
                        reasons.append("MACD金叉，买入信号确认")
                    elif macd.iloc[-1] > 0 and signal.iloc[-1] > 0:
                        reasons.append("MACD处于零轴上方，属于强势区间")
                    
                    # 6. 行业资金分析
                    try:
                        stock_industry = industry_df[industry_df['代码'] == stock['代码']]['板块名称'].iloc[0]
                        industry_flow_data = industry_flow[industry_flow['行业'] == stock_industry]
                        if not industry_flow_data.empty:
                            flow_value = industry_flow_data.iloc[0]['主力净流入-净额']
                            if flow_value > 0:
                                reasons.append(f"{stock_industry}行业资金净流入，板块环境向好")
                            else:
                                reasons.append(f"{stock_industry}行业资金净流出，需要关注板块风险")
                    except:
                        pass
                    
                    # 7. 成交量分析
                    vol_mean = hist_data['成交量'].mean()
                    if latest['成交量'] > vol_mean * 2:
                        reasons.append("成交量显著放大，资金关注度高")
                    elif latest['成交量'] < vol_mean * 0.5:
                        reasons.append("成交量明显萎缩，需要观察量能配合")
                    
                    # 8. 价格位置分析
                    price_max = hist_data['收盘'].max()
                    price_min = hist_data['收盘'].min()
                    price_range = price_max - price_min
                    
                    # 添加边界检查和错误处理
                    if price_range == 0 or price_min == 0:
                        current_position = 50  # 默认值
                    else:
                        current_position = (latest['收盘'] - price_min) / price_range * 100
                    
                    # 计算综合评级分数
                    rating_score = 0
                    risk_score = 0
                    rating_reasons = []
                    
                    # 1. 连续涨停评分
                    if limit_up_count >= 3:
                        rating_score += 2
                        rating_reasons.append("连续三板以上")
                        risk_score += 1  # 同时增加风险分
                    elif limit_up_count == 2:
                        rating_score += 1
                        rating_reasons.append("连续两板")
                    
                    # 2. 量能评分
                    if stock['量比'] > 3:
                        rating_score += 2
                        rating_reasons.append("量能显著放大")
                    elif stock['量比'] > 2:
                        rating_score += 1
                        rating_reasons.append("量能良好")
                    elif stock['量比'] < 0.8:
                        risk_score += 1
                        rating_reasons.append("量能不足")
                    
                    # 3. 均线系统评分
                    if ma5.iloc[-1] > ma10.iloc[-1] > ma20.iloc[-1]:
                        ma5_slope = (ma5.iloc[-1] - ma5.iloc[-5]) / ma5.iloc[-5] * 100
                        if ma5_slope > 2:
                            rating_score += 2
                            rating_reasons.append("均线系统强势")
                        else:
                            rating_score += 1
                            rating_reasons.append("均线系统向好")
                    elif ma5.iloc[-1] < ma10.iloc[-1] < ma20.iloc[-1]:
                        risk_score += 1
                        rating_reasons.append("均线系统弱势")
                    
                    # 4. RSI评分
                    if 50 < rsi.iloc[-1] < 70:
                        rating_score += 1
                        rating_reasons.append("RSI健康")
                    elif rsi.iloc[-1] > 80:
                        risk_score += 2
                        rating_reasons.append("RSI超买")
                    elif rsi.iloc[-1] < 30:
                        risk_score += 1
                        rating_reasons.append("RSI超卖")
                    
                    # 5. MACD评分
                    if macd.iloc[-1] > signal.iloc[-1] and macd.iloc[-2] <= signal.iloc[-2]:
                        rating_score += 2
                        rating_reasons.append("MACD金叉")
                    elif macd.iloc[-1] > 0 and signal.iloc[-1] > 0:
                        rating_score += 1
                        rating_reasons.append("MACD强势")
                    elif macd.iloc[-1] < signal.iloc[-1] and macd.iloc[-2] >= signal.iloc[-2]:
                        risk_score += 1
                        rating_reasons.append("MACD死叉")
                    
                    # 6. 行业资金评分
                    try:
                        stock_industry = industry_df[industry_df['代码'] == stock['代码']]['板块名称'].iloc[0]
                        industry_flow_data = industry_flow[industry_flow['行业'] == stock_industry]
                        if not industry_flow_data.empty:
                            flow_value = industry_flow_data.iloc[0]['主力净流入-净额']
                            if flow_value > 100000000:  # 1亿以上
                                rating_score += 2
                                rating_reasons.append("行业资金大幅流入")
                            elif flow_value > 0:
                                rating_score += 1
                                rating_reasons.append("行业资金净流入")
                            elif flow_value < -100000000:  # -1亿以下
                                risk_score += 2
                                rating_reasons.append("行业资金大幅流出")
                            elif flow_value < 0:
                                risk_score += 1
                                rating_reasons.append("行业资金净流出")
                    except:
                        pass
                    
                    # 7. 价格位置评分
                    if current_position < 30:
                        rating_score += 2
                        rating_reasons.append("低位突破")
                    elif current_position > 70:
                        risk_score += 2
                        rating_reasons.append("高位风险")
                    
                    # 8. 成交量评分
                    vol_mean = hist_data['成交量'].mean()
                    if latest['成交量'] > vol_mean * 3:
                        rating_score += 2
                        rating_reasons.append("成交量显著放大")
                    elif latest['成交量'] > vol_mean * 2:
                        rating_score += 1
                        rating_reasons.append("成交量放大")
                    elif latest['成交量'] < vol_mean * 0.5:
                        risk_score += 1
                        rating_reasons.append("成交量萎缩")
                    
                    # 9. 换手率评分
                    if stock['换手率'] > 15:
                        rating_score += 2
                        rating_reasons.append("换手充分")
                    elif stock['换手率'] > 10:
                        rating_score += 1
                        rating_reasons.append("换手活跃")
                    elif stock['换手率'] < 3:
                        risk_score += 1
                        rating_reasons.append("换手不足")
                    
                    # 生成综合评级
                    rating = ""
                    if rating_score >= 8 and risk_score <= 2:
                        rating = "强烈推荐"
                    elif rating_score >= 6 and risk_score <= 3:
                        rating = "建议关注"
                    elif risk_score >= 5:
                        rating = "强烈风险"
                    elif risk_score >= 3:
                        rating = "注意风险"
                    else:
                        rating = "中性"
                    
                    # 添加到分析结果
                    analysis_results.append({
                        '代码': stock['代码'],
                        '名称': stock['名称'],
                        '现价': stock['最新价'],
                        '涨跌幅': stock['涨跌幅'],
                        '量比': stock['量比'],
                        '换手率': stock['换手率'],
                        '特征': features,
                        '趋势预测': trend_prediction,
                        '原因分析': ' | '.join(reasons),
                        '综合评级': rating,
                        '评级理由': ' | '.join(rating_reasons)
                    })
                    
                except Exception as e:
                    print(f"分析股票 {stock['代码']} 失败: {e}")
                    continue
            
            return analysis_results
            
        except Exception as e:
            print(f"分析涨停票失败: {e}")
            return []

    def predict_trend(self, hist_data, ma5, ma10, ma20, rsi, macd, signal):
        """预测股票趋势"""
        try:
            # 获取最新数据
            latest = hist_data.iloc[-1]
            latest_ma5 = ma5.iloc[-1]
            latest_ma10 = ma10.iloc[-1]
            latest_ma20 = ma20.iloc[-1]
            latest_rsi = rsi.iloc[-1]
            latest_macd = macd.iloc[-1]
            latest_signal = signal.iloc[-1]
            
            # 计算趋势得分
            trend_score = 0
            
            # 1. 均线系统得分
            if latest_ma5 > latest_ma10 > latest_ma20:
                trend_score += 2
            elif latest_ma5 > latest_ma10:
                trend_score += 1
            
            # 2. RSI得分
            if 30 <= latest_rsi <= 70:
                trend_score += 1
            elif latest_rsi > 70:
                trend_score -= 1
            
            # 3. MACD得分
            if latest_macd > latest_signal:
                trend_score += 1
            elif latest_macd < latest_signal:
                trend_score -= 1
            
            # 4. 量能得分
            if latest['成交量'] > hist_data['成交量'].mean():
                trend_score += 1
            
            # 根据得分预测趋势
            if trend_score >= 3:
                return "强势上涨"
            elif trend_score >= 1:
                return "震荡上涨"
            elif trend_score >= -1:
                return "震荡整理"
            else:
                return "可能回调"
            
        except Exception as e:
            print(f"预测趋势失败: {e}")
            return "无法预测"

    def show_limit_up_analysis(self):
        """显示涨停票分析结果"""
        try:
            analysis_results = self.analyze_limit_up_stocks()
            # 清空现有表格内容
            self.result_table.clearContents()
            self.result_table.setRowCount(0)
            columns = ['代码', '名称', '现价', '涨跌幅', '量比', '换手率', '特征', '趋势预测', '原因分析', '综合评级', '评级理由']
            self.result_table.setColumnCount(len(columns))
            self.result_table.setHorizontalHeaderLabels(columns)
            # 填充分析结果
            for result in analysis_results:
                current_row = self.result_table.rowCount()
                self.result_table.insertRow(current_row)
                for col_idx, column in enumerate(columns):
                    item = NumericTableWidgetItem(str(result[column]))
                    item.setTextAlignment(Qt.AlignCenter)
                    # 设置涨跌幅颜色
                    if column == '涨跌幅':
                        try:
                            change = float(str(result[column]).replace('%', ''))
                            if change > 0:
                                item.setForeground(QBrush(QColor('#FF4444')))
                            elif change < 0:
                                item.setForeground(QBrush(QColor('#00AA00')))
                        except:
                            pass
                    # 设置趋势预测颜色
                    if column == '趋势预测':
                        if "上涨" in result[column]:
                            item.setForeground(QBrush(QColor('#FF4444')))
                        elif "回调" in result[column]:
                            item.setForeground(QBrush(QColor('#00AA00')))
                        else:
                            item.setForeground(QBrush(QColor('#0088FF')))
                    # 设置原因分析颜色
                    if column == '原因分析':
                        if "风险" in result[column]:
                            item.setForeground(QBrush(QColor('#FF4444')))
                        elif "向好" in result[column] or "强劲" in result[column]:
                            item.setForeground(QBrush(QColor('#00AA00')))
                    # 设置综合评级颜色
                    if column == '综合评级':
                        if "强烈推荐" in result[column]:
                            item.setForeground(QBrush(QColor('#FF0000')))
                            item.setBackground(QBrush(QColor('#FFEEEE')))
                        elif "建议关注" in result[column]:
                            item.setForeground(QBrush(QColor('#FF4444')))
                        elif "强烈风险" in result[column]:
                            item.setForeground(QBrush(QColor('#FFFFFF')))
                            item.setBackground(QBrush(QColor('#FF4444')))
                        elif "注意风险" in result[column]:
                            item.setForeground(QBrush(QColor('#FF0000')))
                    # 设置评级理由颜色
                    if column == '评级理由':
                        if "风险" in result[column] or "不足" in result[column] or "弱势" in result[column]:
                            item.setForeground(QBrush(QColor('#FF4444')))
                        elif "强势" in result[column] or "放大" in result[column] or "突破" in result[column]:
                            item.setForeground(QBrush(QColor('#00AA00')))
                    self.result_table.setItem(current_row, col_idx, item)
            # 添加导出Excel按钮
            export_btn = QPushButton("导出Excel")
            export_btn.clicked.connect(lambda: self.export_to_excel(self.result_table))
            # 在表格下方弹窗显示
            msg_box = QDialog(self)
            msg_box.setWindowTitle("涨停票分析结果")
            layout = QVBoxLayout()
            layout.addWidget(self.result_table)
            btn_layout = QHBoxLayout()
            btn_layout.addWidget(export_btn)
            close_btn = QPushButton("关闭")
            close_btn.clicked.connect(msg_box.close)
            btn_layout.addWidget(close_btn)
            layout.addLayout(btn_layout)
            msg_box.setLayout(layout)
            msg_box.exec_()
        except Exception as e:
            print(f"显示涨停票分析失败: {e}")

    def analyze_market_trend(self):
        """分析大盘趋势"""
        try:
            # 获取上证指数数据
            sh_index = ak.stock_zh_index_daily_em(symbol="sh000001")
            if sh_index.empty:
                print("获取上证指数数据为空")
                return None
                
            # 打印列名，用于调试
            print("上证指数数据列名:", sh_index.columns.tolist())
            
            # 检查并重命名列名（如果需要）
            column_mapping = {
                'close': '收盘',
                'open': '开盘',
                'high': '最高',
                'low': '最低',
                'volume': '成交量'
            }
            
            # 重命名列名
            for old_name, new_name in column_mapping.items():
                if old_name in sh_index.columns and new_name not in sh_index.columns:
                    sh_index = sh_index.rename(columns={old_name: new_name})
            
            # 再次打印列名，确认重命名后的结果
            print("重命名后的列名:", sh_index.columns.tolist())
                
            # 计算技术指标
            ma5 = sh_index['收盘'].rolling(5).mean()
            ma10 = sh_index['收盘'].rolling(10).mean()
            ma20 = sh_index['收盘'].rolling(20).mean()
            ma60 = sh_index['收盘'].rolling(60).mean()
            
            # 计算MACD
            exp1 = sh_index['收盘'].ewm(span=12, adjust=False).mean()
            exp2 = sh_index['收盘'].ewm(span=26, adjust=False).mean()
            macd = exp1 - exp2
            signal = macd.ewm(span=9, adjust=False).mean()
            
            # 计算RSI
            delta = sh_index['收盘'].diff()
            gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
            loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
            rs = gain / loss
            rsi = 100 - (100 / (1 + rs))
            
            # 获取最新数据
            latest = sh_index.iloc[-1]
            latest_ma5 = ma5.iloc[-1]
            latest_ma10 = ma10.iloc[-1]
            latest_ma20 = ma20.iloc[-1]
            latest_ma60 = ma60.iloc[-1]
            latest_macd = macd.iloc[-1]
            latest_signal = signal.iloc[-1]
            latest_rsi = rsi.iloc[-1]
            
            # 分析趋势
            trend_analysis = {
                '均线系统': {
                    '多头排列': latest_ma5 > latest_ma10 > latest_ma20 > latest_ma60,
                    '空头排列': latest_ma5 < latest_ma10 < latest_ma20 < latest_ma60,
                    'MA5斜率': (latest_ma5 - ma5.iloc[-5]) / ma5.iloc[-5] * 100,
                    'MA10斜率': (latest_ma10 - ma10.iloc[-5]) / ma10.iloc[-5] * 100
                },
                'MACD': {
                    '金叉': latest_macd > latest_signal and macd.iloc[-2] <= signal.iloc[-2],
                    '死叉': latest_macd < latest_signal and macd.iloc[-2] >= signal.iloc[-2],
                    '零轴上方': latest_macd > 0 and latest_signal > 0,
                    '零轴下方': latest_macd < 0 and latest_signal < 0
                },
                'RSI': {
                    '超买': latest_rsi > 80,
                    '超卖': latest_rsi < 20,
                    '强势': 50 < latest_rsi < 70,
                    '弱势': 30 < latest_rsi < 50
                }
            }
            
            # 计算支撑压力位
            recent_high = sh_index['最高'].tail(20).max()
            recent_low = sh_index['最低'].tail(20).min()
            current_price = latest['收盘']
            
            # 计算资金流向
            volume_ma5 = sh_index['成交量'].rolling(5).mean()
            volume_ma10 = sh_index['成交量'].rolling(10).mean()
            
            # 判断趋势
            trend = "震荡整理"
            if trend_analysis['均线系统']['多头排列'] and trend_analysis['MACD']['金叉']:
                trend = "强势上涨"
            elif trend_analysis['均线系统']['多头排列'] and trend_analysis['MACD']['零轴上方']:
                trend = "震荡上涨"
            elif trend_analysis['均线系统']['空头排列'] and trend_analysis['MACD']['死叉']:
                trend = "强势下跌"
            elif trend_analysis['均线系统']['空头排列'] and trend_analysis['MACD']['零轴下方']:
                trend = "震荡下跌"
            
            # 生成建议
            advice = ""
            if trend == "强势上涨":
                advice = "可以积极做多，但注意高位风险"
            elif trend == "震荡上涨":
                advice = "可以逢低买入，注意节奏"
            elif trend == "强势下跌":
                advice = "建议清仓观望，等待企稳"
            elif trend == "震荡下跌":
                advice = "建议轻仓观望，等待企稳"
            else:
                advice = "建议观望，等待方向明确"

            # 分析板块情况
            try:
                # 获取行业板块数据
                industry_df = ak.stock_board_industry_name_em()
                # 获取行业资金流向
                industry_flow = ak.stock_sector_fund_flow_rank()
                
                # 打印行业资金流向数据的列名和示例数据
                print("行业资金流向数据列名:", industry_flow.columns.tolist())
                print("行业资金流向数据示例:\n", industry_flow.head())
                
                # 分析强势板块
                strong_sectors = []
                potential_sectors = []
                
                # 获取行业涨跌幅
                industry_change = ak.stock_board_industry_cons_em()

                # 添加字符串转数字的辅助函数
                def convert_flow_value(value):
                    try:
                        if isinstance(value, (int, float)):
                            return float(value)
                        if not isinstance(value, str):
                            return 0
                        # 去除空格和特殊字符
                        value = value.strip().replace(',', '')
                        
                        # 处理负数
                        is_negative = value.startswith('-')
                        if is_negative:
                            value = value[1:]
                        
                        # 处理单位
                        if '亿' in value:
                            base_value = float(value.replace('亿', ''))
                            result = base_value * 100000000
                        elif '万' in value:
                            base_value = float(value.replace('万', ''))
                            result = base_value * 10000
                        else:
                            result = float(value)
                        
                        # 恢复负号
                        return -result if is_negative else result
                        
                    except Exception as e:
                        print(f"转换资金值失败: {value}, 错误: {e}")
                        return 0
                
                # 分析每个行业板块
                for industry in industry_df['板块名称'].unique():
                    try:
                        # 获取该行业的股票
                        industry_stocks = ak.stock_board_industry_cons_em(symbol=industry)
                        if industry_stocks.empty:
                            continue
                            
                        # 计算行业平均涨跌幅
                        avg_change = industry_stocks['涨跌幅'].mean()
                        
                        # 获取行业资金流向
                        industry_flow_data = industry_flow[industry_flow['名称'] == industry]
                        if not industry_flow_data.empty:
                            # 打印当前行业的资金流向数据
                            print(f"\n处理行业: {industry}")
                            print("行业资金流向数据:\n", industry_flow_data)
                            
                            # 尝试不同的字段名
                            flow_value = 0
                            possible_fields = ['今日主力净流入-净额', '主力净流入-净额', '主力净额']
                            for field in possible_fields:
                                if field in industry_flow_data.columns:
                                    value = industry_flow_data.iloc[0][field]
                                    print(f"尝试字段 {field}: {value}")
                                    flow_value = convert_flow_value(value)
                                    if flow_value != 0:
                                        break
                            
                            # 判断强势板块
                            if avg_change > 2 and flow_value > 0:
                                strong_sectors.append({
                                    '名称': industry,
                                    '涨跌幅': f"{avg_change:.2f}%",
                                    '资金流入': f"{flow_value/100000000:.2f}亿",
                                    '资金流入值': flow_value  # 添加原始数值
                                })
                            
                            # 判断潜力板块
                            if avg_change < 2 and flow_value > 50000000:  # 5000万以上
                                potential_sectors.append({
                                    '名称': industry,
                                    '涨跌幅': f"{avg_change:.2f}%",
                                    '资金流入': f"{flow_value/100000000:.2f}亿",
                                    '资金流入值': flow_value  # 添加原始数值
                                })
                    except Exception as e:
                        print(f"分析行业 {industry} 时出错: {e}")
                        continue
                
                # 按涨跌幅排序
                strong_sectors.sort(key=lambda x: float(x['涨跌幅'].replace('%', '')), reverse=True)
                # 使用原始数值进行排序
                potential_sectors.sort(key=lambda x: x['资金流入值'], reverse=True)
                
            except Exception as e:
                print(f"分析板块情况时出错: {e}")
                import traceback
                traceback.print_exc()
                strong_sectors = []
                potential_sectors = []
            
            return {
                '趋势': trend,
                '建议': advice,
                '支撑位': recent_low,
                '压力位': recent_high,
                '当前价格': current_price,
                '技术分析': trend_analysis,
                '成交量分析': {
                    '放量': latest['成交量'] > volume_ma5.iloc[-1] * 1.5,
                    '缩量': latest['成交量'] < volume_ma5.iloc[-1] * 0.8,
                    '量能趋势': "上升" if volume_ma5.iloc[-1] > volume_ma10.iloc[-1] else "下降"
                },
                '强势板块': strong_sectors,
                '潜力板块': potential_sectors
            }
            
        except Exception as e:
            print(f"分析大盘趋势失败: {e}")
            import traceback
            traceback.print_exc()
            return None

    def analyze_money_flow(self):
        """分析资金流向"""
        try:
            # 获取个股资金流向数据
            stock_flow = ak.stock_individual_fund_flow_rank()
            print("\n=== 数据结构检查 ===")
            print("列名:", stock_flow.columns.tolist())
            print("数据示例:\n", stock_flow.head())
            
            # 获取A股实时行情数据以获取股票名称
            stock_info = ak.stock_zh_a_spot_em()
            stock_name_map = dict(zip(stock_info['代码'], stock_info['名称']))
            
            # 处理个股资金流向
            print("\n=== 处理个股资金流向 ===")
            stock_flows_5000w = []  # 5000万以上
            stock_flows_1000w = []  # 1000万-5000万
            stock_flows_100w = []   # 100万-1000万
            
            for _, row in stock_flow.iterrows():
                try:
                    # 获取股票代码
                    stock_code = row['代码'] if '代码' in row else None
                    if not stock_code:
                        print(f"找不到股票代码，行数据: {row}")
                        continue
                        
                    stock_name = stock_name_map.get(stock_code, f"未知股票{stock_code}")
                    
                    # 获取主力净流入
                    main_flow = 0
                    for col in row.index:
                        if '主力净流入' in col:
                            try:
                                value = str(row[col]).replace(',', '')
                                if '亿' in value:
                                    main_flow = float(value.replace('亿', '')) * 100000000
                                elif '万' in value:
                                    main_flow = float(value.replace('万', '')) * 10000
                                else:
                                    main_flow = float(value)
                                break
                            except:
                                continue
                    
                    # 获取超大单净流入
                    super_flow = 0
                    for col in row.index:
                        if '超大单净流入' in col:
                            try:
                                value = str(row[col]).replace(',', '')
                                if '亿' in value:
                                    super_flow = float(value.replace('亿', '')) * 100000000
                                elif '万' in value:
                                    super_flow = float(value.replace('万', '')) * 10000
                                else:
                                    super_flow = float(value)
                                break
                            except:
                                continue
                    
                    # 获取大单净流入
                    big_flow = 0
                    for col in row.index:
                        if '大单净流入' in col and '超' not in col:
                            try:
                                value = str(row[col]).replace(',', '')
                                if '亿' in value:
                                    big_flow = float(value.replace('亿', '')) * 100000000
                                elif '万' in value:
                                    big_flow = float(value.replace('万', '')) * 10000
                                else:
                                    big_flow = float(value)
                                break
                            except:
                                continue
                    
                    # 获取中单净流入
                    mid_flow = 0
                    for col in row.index:
                        if '中单净流入' in col:
                            try:
                                value = str(row[col]).replace(',', '')
                                if '亿' in value:
                                    mid_flow = float(value.replace('亿', '')) * 100000000
                                elif '万' in value:
                                    mid_flow = float(value.replace('万', '')) * 10000
                                else:
                                    mid_flow = float(value)
                                break
                            except:
                                continue
                    
                    # 获取涨跌幅
                    change_pct = 0
                    for col in row.index:
                        if '涨跌幅' in col:
                            try:
                                value = str(row[col]).replace('%', '')
                                change_pct = float(value)
                                break
                            except:
                                continue
                    
                    # 创建股票信息
                    stock_info = {
                        'code': stock_code,
                        'name': stock_name,
                        'change_pct': change_pct,
                        'flow': main_flow,
                        '超大单': super_flow,
                        '大单': big_flow,
                        '中单': mid_flow
                    }
                    
                    # 根据资金流向大小分类
                    if abs(main_flow) >= 50000000:  # 5000万
                        stock_flows_5000w.append(stock_info)
                    elif abs(main_flow) >= 10000000:  # 1000万
                        stock_flows_1000w.append(stock_info)
                    elif abs(main_flow) >= 1000000:  # 100万
                        stock_flows_100w.append(stock_info)
                        
                except Exception as e:
                    print(f"处理行数据时出错: {e}")
                    print(f"错误行数据: {row}")
                    continue
            
            # 按资金流向绝对值排序
            stock_flows_5000w.sort(key=lambda x: abs(x['flow']), reverse=True)
            stock_flows_1000w.sort(key=lambda x: abs(x['flow']), reverse=True)
            stock_flows_100w.sort(key=lambda x: abs(x['flow']), reverse=True)
            
            return {
                'stock_flows_5000w': stock_flows_5000w,
                'stock_flows_1000w': stock_flows_1000w,
                'stock_flows_100w': stock_flows_100w
            }
            
        except Exception as e:
            print(f"分析资金流向时出错: {e}")
            import traceback
            traceback.print_exc()
            return None

    def show_market_analysis(self):
        """显示市场分析结果"""
        try:
            # 获取资金流向分析
            money_flow = self.analyze_money_flow()
            if not money_flow:
                return
            
            # 创建对话框
            dialog = QDialog(self)
            dialog.setWindowTitle("市场资金流向分析")
            dialog.setMinimumWidth(800)
            
            # 创建主布局
            main_layout = QVBoxLayout()
            
            # 创建资金流向分组
            flow_group = QGroupBox("个股资金流向分析")
            flow_layout = QVBoxLayout()
            
            # 显示5000万以上个股
            if money_flow['stock_flows_5000w']:
                flow_layout.addWidget(QLabel("\n【资金流入5000万以上】"))
                for stock in money_flow['stock_flows_5000w']:
                    flow_text = f"{stock['name']}({stock['code']}) "
                    flow_text += f"涨跌幅: {stock['change_pct']}%, "
                    flow_text += f"主力净流入: {stock['flow']/100000000:.2f}亿, "
                    flow_text += f"超大单: {stock['超大单']/100000000:.2f}亿, "
                    flow_text += f"大单: {stock['大单']/100000000:.2f}亿, "
                    flow_text += f"中单: {stock['中单']/100000000:.2f}亿, "
                    flow_text += f"小单: {stock['小单']/100000000:.2f}亿"
                    flow_layout.addWidget(QLabel(flow_text))
            
            # 显示1000万-5000万个股
            if money_flow['stock_flows_1000w']:
                flow_layout.addWidget(QLabel("\n【资金流入1000万-5000万】"))
                for stock in money_flow['stock_flows_1000w']:
                    flow_text = f"{stock['name']}({stock['code']}) "
                    flow_text += f"涨跌幅: {stock['change_pct']}%, "
                    flow_text += f"主力净流入: {stock['flow']/100000000:.2f}亿, "
                    flow_text += f"超大单: {stock['超大单']/100000000:.2f}亿, "
                    flow_text += f"大单: {stock['大单']/100000000:.2f}亿, "
                    flow_text += f"中单: {stock['中单']/100000000:.2f}亿, "
                    flow_text += f"小单: {stock['小单']/100000000:.2f}亿"
                    flow_layout.addWidget(QLabel(flow_text))
            
            # 显示100万-1000万个股
            if money_flow['stock_flows_100w']:
                flow_layout.addWidget(QLabel("\n【资金流入100万-1000万】"))
                for stock in money_flow['stock_flows_100w']:
                    flow_text = f"{stock['name']}({stock['code']}) "
                    flow_text += f"涨跌幅: {stock['change_pct']}%, "
                    flow_text += f"主力净流入: {stock['flow']/100000000:.2f}亿, "
                    flow_text += f"超大单: {stock['超大单']/100000000:.2f}亿, "
                    flow_text += f"大单: {stock['大单']/100000000:.2f}亿, "
                    flow_text += f"中单: {stock['中单']/100000000:.2f}亿, "
                    flow_text += f"小单: {stock['小单']/100000000:.2f}亿"
                    flow_layout.addWidget(QLabel(flow_text))
            
            flow_group.setLayout(flow_layout)
            main_layout.addWidget(flow_group)
            
            # 设置对话框布局
            dialog.setLayout(main_layout)
            dialog.exec_()
            
        except Exception as e:
            print(f"显示市场分析时出错: {e}")
            import traceback
            traceback.print_exc()

    def show_money_flow_analysis(self):
        """显示资金流向分析结果"""
        try:
            money_flow = self.analyze_money_flow()
            if not money_flow:
                return
            dialog = QDialog(self)
            dialog.setWindowTitle("资金流向分析")
            dialog.setMinimumWidth(1000)
            main_layout = QVBoxLayout()
            # 创建表格
            table = QTableWidget()
            table.setColumnCount(7)
            table.setHorizontalHeaderLabels(['代码', '名称', '涨跌幅', '主力净流入', '超大单', '大单', '中单'])
            all_data = money_flow['stock_flows_5000w'] + money_flow['stock_flows_1000w'] + money_flow['stock_flows_100w']
            table.setRowCount(len(all_data))
            for i, stock in enumerate(all_data):
                table.setItem(i, 0, QTableWidgetItem(stock['code']))
                table.setItem(i, 1, QTableWidgetItem(stock['name']))
                table.setItem(i, 2, NumericTableWidgetItem(str(stock['change_pct'])))
                table.setItem(i, 3, NumericTableWidgetItem(str(stock['flow'])))
                table.setItem(i, 4, NumericTableWidgetItem(str(stock['超大单'])))
                table.setItem(i, 5, NumericTableWidgetItem(str(stock['大单'])))
                table.setItem(i, 6, NumericTableWidgetItem(str(stock['中单'])))
            main_layout.addWidget(table)
            # 添加导出Excel按钮
            export_btn = QPushButton("导出Excel")
            export_btn.clicked.connect(lambda: self.export_to_excel(table))
            btn_layout = QHBoxLayout()
            btn_layout.addWidget(export_btn)
            close_btn = QPushButton("关闭")
            close_btn.clicked.connect(dialog.close)
            btn_layout.addWidget(close_btn)
            main_layout.addLayout(btn_layout)
            dialog.setLayout(main_layout)
            dialog.exec_()
        except Exception as e:
            print(f"显示资金流向分析时出错: {e}")

    def show_money_flow_rank(self):
        """显示资金净流入排序"""
        try:
            import akshare as ak
            rank_window = QDialog(self)
            rank_window.setWindowTitle("个股资金净流入排序")
            rank_window.setMinimumWidth(1400)
            
            # 创建主布局
            layout = QVBoxLayout()
            
            # 创建时间范围选择下拉框
            time_range_combo = QComboBox()
            time_range_combo.addItems(["东方财富当日", "当日", "5日"])
            
            # 创建刷新按钮和导出按钮
            refresh_btn = QPushButton("刷新")
            export_btn = QPushButton("导出Excel")
            refresh_btn.clicked.connect(lambda: update_table(time_range_combo.currentText()))
            
            # 创建水平布局来放置下拉框和按钮
            top_layout = QHBoxLayout()
            top_layout.addWidget(QLabel("时间范围:"))
            top_layout.addWidget(time_range_combo)
            top_layout.addWidget(refresh_btn)
            top_layout.addWidget(export_btn)
            top_layout.addStretch()
            
            # 将顶部布局添加到主布局
            layout.addLayout(top_layout)
            
            # 创建表格
            flow_table = QTableWidget()
            flow_table.setColumnCount(8)
            flow_table.setHorizontalHeaderLabels([
                '股票代码', '股票名称', '最新价', '涨跌幅', 
                '主力净流入(亿)', '超大单净流入(亿)', '大单净流入(亿)', '中单净流入(亿)'
            ])
            layout.addWidget(flow_table)
            
            # 创建关闭按钮
            close_btn = QPushButton("关闭")
            close_btn.clicked.connect(rank_window.close)
            layout.addWidget(close_btn)
            
            # 设置对话框布局
            rank_window.setLayout(layout)
            
            def update_table(time_range):
                try:
                    print(f"正在获取{time_range}资金流向数据...")
                    if time_range == "当日":
                        flow_data = ak.stock_fund_flow_individual()
                        column_map = {
                            'code': '股票代码',
                            'name': '股票简称',
                            'price': '最新价',
                            'change_percent': '涨跌幅',
                            'main_net_inflow': '净额',
                            'super_net_inflow': '流入资金',
                            'big_net_inflow': '流出资金',
                            'medium_net_inflow': '成交额'
                        }
                    elif time_range == "5日":
                        flow_data = ak.stock_individual_fund_flow_rank()
                        column_map = {
                            'code': '代码',
                            'name': '名称',
                            'price': '最新价',
                            'change_percent': '5日涨跌幅',
                            'main_net_inflow': '5日主力净流入-净额',
                            'super_net_inflow': '5日超大单净流入-净额',
                            'big_net_inflow': '5日大单净流入-净额',
                            'medium_net_inflow': '5日中单净流入-净额'
                        }
                    elif time_range == "东方财富当日":
                        flow_data = ak.stock_individual_fund_flow_rank(indicator="今日")
                        # 字段适配
                        column_map = {
                            'code': '代码',
                            'name': '名称',
                            'price': '最新价',
                            'change_percent': '今日涨跌幅',
                            'main_net_inflow': '今日主力净流入-净额',
                            'super_net_inflow': '今日超大单净流入-净额',
                            'big_net_inflow': '今日大单净流入-净额',
                            'medium_net_inflow': '今日中单净流入-净额'
                        }
                    else:
                        return
                    print(f"数据形状: {flow_data.shape}")
                    print(f"数据列名: {flow_data.columns.tolist()}")
                    sorted_data = []
                    flow_table.setRowCount(len(flow_data))
                    for i, row in flow_data.iterrows():
                        try:
                            code = str(row[column_map['code']])
                            name = str(row[column_map['name']])
                            price = str(row[column_map['price']])
                            change = str(row[column_map['change_percent']]).rstrip('%') + '%'
                            flow_table.setItem(i, 0, QTableWidgetItem(code))
                            flow_table.setItem(i, 1, QTableWidgetItem(name))
                            flow_table.setItem(i, 2, NumericTableWidgetItem(price))
                            flow_table.setItem(i, 3, NumericTableWidgetItem(change))
                            def convert_flow_value(value):
                                try:
                                    if isinstance(value, str):
                                        value = value.replace(',', '')
                                        if '亿' in value:
                                            return float(value.replace('亿', ''))
                                        elif '万' in value:
                                            return float(value.replace('万', '')) / 10000
                                        else:
                                            return float(value) / 100000000
                                    else:
                                        return float(value) / 100000000
                                except:
                                    return 0.0
                            flow_columns = [
                                column_map['main_net_inflow'],
                                column_map['super_net_inflow'],
                                column_map['big_net_inflow'],
                                column_map['medium_net_inflow']
                            ]
                            flow_values = []
                            for column_name in flow_columns:
                                value = convert_flow_value(row[column_name])
                                flow_values.append(value)
                            sorted_data.append({
                                'code': code,
                                'name': name,
                                'price': price,
                                'change': change,
                                'flows': flow_values,
                                'main_flow': flow_values[0]
                            })
                            for col, value in enumerate(flow_values, start=4):
                                item = NumericTableWidgetItem(f"{value:.2f}")
                                if value > 0:
                                    item.setForeground(QBrush(QColor('#FF4444')))
                                else:
                                    item.setForeground(QBrush(QColor('#00AA00')))
                                flow_table.setItem(i, col, item)
                            try:
                                change_item = flow_table.item(i, 3)
                                change_text = change_item.text().replace('%', '')
                                change_value = float(change_text)
                                if change_value > 0:
                                    change_item.setForeground(QBrush(QColor('#FF4444')))
                                elif change_value < 0:
                                    change_item.setForeground(QBrush(QColor('#00AA00')))
                            except Exception as e:
                                print(f"设置涨跌幅颜色失败: {e}")
                        except Exception as e:
                            print(f"处理第 {i+1} 条数据失败: {e}")
                            continue
                    sorted_data.sort(key=lambda x: x['main_flow'], reverse=True)
                    for data in sorted_data:
                        flow_str = " ".join([f"{flow:>12.2f}" for flow in data['flows']])
                        print("{:<10} {:<12} {:<8} {:<8} {}".format(
                            data['code'],
                            data['name'],
                            data['price'],
                            data['change'],
                            flow_str
                        ))
                    flow_table.setSortingEnabled(True)
                    flow_table.resizeColumnsToContents()
                    flow_table.setAlternatingRowColors(True)
                except Exception as e:
                    print(f"获取资金流向数据失败: {e}")
                    QMessageBox.critical(rank_window, "错误", "获取资金流向数据失败，请稍后重试")
            
            # 连接信号
            time_range_combo.currentTextChanged.connect(update_table)
            export_btn.clicked.connect(lambda: self.export_to_excel(flow_table))
            
            # 初始化显示
            update_table("东方财富当日")
            
            # 显示对话框
            rank_window.exec_()
            
        except Exception as e:
            print(f"显示资金净流入排序时出错: {e}")
            QMessageBox.critical(self, "错误", "显示资金净流入排序失败，请稍后重试")

    def show_context_menu(self, pos):
        # 获取当前表格
        table = self.sender()
        if not table:
            return

        # 创建右键菜单
        menu = QMenu(self)
        export_action = QAction("导出为Excel", self)
        export_action.triggered.connect(lambda: self.export_to_excel(table))
        menu.addAction(export_action)
        
        # 显示菜单
        menu.exec_(table.mapToGlobal(pos))

    def export_to_excel(self, table):
        try:
            # 获取当前时间作为文件名
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # 创建保存文件的对话框
            file_name, _ = QFileDialog.getSaveFileName(
                self,
                "保存Excel文件",
                f"stock_data_{current_time}.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if not file_name:
                return
                
            # 获取表格数据
            data = []
            headers = []
            
            # 获取表头
            for col in range(table.columnCount()):
                headers.append(table.horizontalHeaderItem(col).text())
            
            # 获取数据
            for row in range(table.rowCount()):
                row_data = []
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    if item is not None:
                        row_data.append(item.text())
                    else:
                        row_data.append("")
                data.append(row_data)
            
            # 创建DataFrame
            df = pd.DataFrame(data, columns=headers)
            
            # 保存为Excel
            df.to_excel(file_name, index=False)
            
            QMessageBox.information(self, "成功", "数据已成功导出到Excel文件！")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出Excel时发生错误：{str(e)}")

    def show_main_fund_rank(self):
        """显示主力排名表格"""
        import akshare as ak
        import pandas as pd
        try:
            dialog = QDialog(self)
            dialog.setWindowTitle("主力排名")
            dialog.setMinimumWidth(1200)
            
            # 创建主布局
            layout = QVBoxLayout()
            
            # 创建顶部按钮布局
            top_layout = QHBoxLayout()
            
            # 创建刷新按钮
            refresh_btn = QPushButton("刷新数据")
            refresh_btn.clicked.connect(lambda: update_table())
            
            # 创建导出按钮
            export_btn = QPushButton("导出Excel")
            
            # 添加按钮到顶部布局
            top_layout.addWidget(refresh_btn)
            top_layout.addWidget(export_btn)
            top_layout.addStretch()
            
            # 将顶部布局添加到主布局
            layout.addLayout(top_layout)
            
            # 创建表格
            table = QTableWidget()
            layout.addWidget(table)
            
            def update_table():
                try:
                    # 获取数据
                    df = ak.stock_main_fund_flow()
                    
                    # 设置表格
                    columns = df.columns.tolist()
                    table.setColumnCount(len(columns))
                    table.setHorizontalHeaderLabels(columns)
                    table.setRowCount(len(df))
                    
                    # 填充数据
                    for i, row in df.iterrows():
                        for j, col in enumerate(columns):
                            #NumericTableWidgetItem QTableWidgetItem
                            item = NumericTableWidgetItem(str(row[col]))
                            item.setTextAlignment(Qt.AlignCenter)
                            
                            # 设置涨跌幅颜色
                            if '涨跌幅' in col or '涨跌幅(%)' in col:
                                try:
                                    value = float(str(row[col]).replace('%', ''))
                                    if value > 0:
                                        item.setForeground(QBrush(QColor('#FF4444')))
                                    elif value < 0:
                                        item.setForeground(QBrush(QColor('#00AA00')))
                                except:
                                    pass
                            
                            # 设置资金流向颜色
                            if '净额' in col or '净流入' in col:
                                try:
                                    value = float(str(row[col]).replace(',', ''))
                                    if value > 0:
                                        item.setForeground(QBrush(QColor('#FF4444')))
                                    elif value < 0:
                                        item.setForeground(QBrush(QColor('#00AA00')))
                                except:
                                    pass
                            
                            table.setItem(i, j, item)
                    
                    # 设置表格属性
                    table.setSortingEnabled(True)
                    table.resizeColumnsToContents()
                    table.setAlternatingRowColors(True)
                    
                except Exception as e:
                    print(f"更新主力排名数据失败: {e}")
                    QMessageBox.critical(dialog, "错误", "获取主力排名数据失败，请稍后重试")
            
            # 创建底部按钮布局
            bottom_layout = QHBoxLayout()
            close_btn = QPushButton("关闭")
            close_btn.clicked.connect(dialog.close)
            bottom_layout.addWidget(close_btn)
            
            # 将底部布局添加到主布局
            layout.addLayout(bottom_layout)
            
            # 设置对话框布局
            dialog.setLayout(layout)
            
            # 连接导出按钮
            export_btn.clicked.connect(lambda: self.export_to_excel(table))
            
            # 初始化显示数据
            update_table()
            
            # 显示对话框
            dialog.exec_()
            
        except Exception as e:
            print(f"显示主力排名失败: {e}")
            QMessageBox.critical(self, "错误", f"获取主力排名数据失败: {e}")


def main():
    app = QApplication(sys.argv)
    screener = StockScreener()
    screener.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
