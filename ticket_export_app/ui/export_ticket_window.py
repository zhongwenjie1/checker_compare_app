# -*- coding: utf-8 -*-
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QToolBar, QStatusBar, QMessageBox, QTableWidget,
    QTableWidgetItem, QSpinBox, QComboBox, QLineEdit, QColorDialog,
    QTabWidget, QFrame, QAbstractItemView, QHeaderView, QGraphicsScene, QGraphicsView, QProgressBar,
    QPlainTextEdit
)
from PySide6.QtCore import Qt, QThreadPool, QTimer
from PySide6.QtGui import QAction, QColor, QPen, QBrush
import os
import math
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill, Border, Side

# 版本号：优先从当前工程的 __init__ 里取，取不到就用 "dev"
try:
    from __init__ import __version__
except Exception:
    __version__ = "dev"

# Worker & tickets：直接从当前工程内部模块导入
from infra.threads import Worker
from core import tickets


class ExportTicketWindow(QMainWindow):
    """
    导出组合票（独立于数据校对）
    v2 岗位矩阵字段：
      序号 / 工程名称 / 设备数量 / 所属线别 / 岗位设备 / A工时 / B工时 / C工时
    说明：
      - 设备数量：1 表示单资源；2 表示双线双资源。
      - 所属线别：1号线 / 2号线 / 双线 / 双线共用。
      - A/B/C 工时 > 0：该车型经过该岗位；工时 = 0：该车型跳过该岗位。
      - 参与投车的车型，工时不能为空；未参与投车的车型，工时可以为空。
    """
    COL_C_TIME = 7
    MAX_SINGLE_STEPS = 23  # 单工程组合票：新版模板固定支持 23 行

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"导出组合票  v{__version__}")
        self.resize(1120, 720)

        self.thread_pool = QThreadPool.globalInstance()
        self.dst_path = None

        self._build_ui()
        self._connect_signals()
    def _on_tab_changed(self, index: int):
        """
        Tab 切换时，控制多工程页内“添加步骤 / 删除步骤 / 填入示例”按钮：
        - 仅在『多工程组合票』页签（第 0 个 Tab）启用；
        - 在『单工程组合票』页签禁用，避免误点影响单工程表。
        """
        is_multi = (index == 0)
        if hasattr(self, "btn_add_row"):
            self.btn_add_row.setEnabled(is_multi)
        if hasattr(self, "btn_del_row"):
            self.btn_del_row.setEnabled(is_multi)
        if hasattr(self, "btn_fill_sample"):
            self.btn_fill_sample.setEnabled(is_multi)
    # ---------------- UI ---------------- #
    def _build_ui(self):
        tb = QToolBar("Ticket")
        self.addToolBar(tb)

        self.act_help = QAction("帮助", self)
        tb.addAction(self.act_help)
        tb.addSeparator()

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # ====== Tab 控件 ======
        self.tabs = QTabWidget(self)
        root.addWidget(self.tabs)

        # ---------- Tab1：多车组合票 ----------
        self.page_multi = QWidget(self)
        page_multi_layout = QVBoxLayout(self.page_multi)
        page_multi_layout.setContentsMargins(8, 8, 8, 8)
        page_multi_layout.setSpacing(10)
        self.tabs.addTab(self.page_multi, "多工程组合票")

        # 多工程内部二级页：第一页负责录入，第二页负责分析/导出/后续动画预留
        self.multi_tabs = QTabWidget(self.page_multi)
        page_multi_layout.addWidget(self.multi_tabs, 1)

        self.page_multi_input = QWidget(self.page_multi)
        self.page_multi_result = QWidget(self.page_multi)

        input_layout = QVBoxLayout(self.page_multi_input)
        input_layout.setContentsMargins(0, 0, 0, 0)
        input_layout.setSpacing(10)

        result_layout = QVBoxLayout(self.page_multi_result)
        result_layout.setContentsMargins(0, 0, 0, 0)
        result_layout.setSpacing(10)


        self.multi_tabs.addTab(self.page_multi_input, "参数与岗位")
        self.multi_tabs.addTab(self.page_multi_result, "分析与导出")


        def _make_block(title: str):
            frame = QFrame(self.page_multi)
            frame.setFrameShape(QFrame.StyledPanel)
            frame.setObjectName("ticketBlock")
            frame.setStyleSheet(
                "QFrame#ticketBlock {"
                "background: #ffffff;"
                "border: 1px solid #d9dee7;"
                "border-radius: 10px;"
                "}"
            )
            lay = QVBoxLayout(frame)
            lay.setContentsMargins(12, 12, 12, 12)
            lay.setSpacing(10)
            title_label = QLabel(title, frame)
            title_label.setStyleSheet("font-size: 14px; font-weight: 700; color: #223042;")
            lay.addWidget(title_label)
            return frame, lay

        top_split = QVBoxLayout()
        top_split.setSpacing(10)
        input_layout.addLayout(top_split)

        # 左侧：运行参数区块
        params_frame, params_layout = _make_block("运行参数")
        top_split.addWidget(params_frame, 0)

        row_top_1 = QHBoxLayout()
        row_top_1.setSpacing(8)
        params_layout.addLayout(row_top_1)
        row_top_1.addWidget(QLabel("工程名称："))
        self.ed_project = QLineEdit()
        self.ed_project.setPlaceholderText("例如：L2++")
        self.ed_project.setFixedWidth(220)
        row_top_1.addWidget(self.ed_project)

        row_top_1.addSpacing(16)
        row_top_1.addWidget(QLabel("投车模式："))
        self.cmb_launch_mode = QComboBox()
        self.cmb_launch_mode.addItems(["按数量投车", "按比例投车"])
        self.cmb_launch_mode.setCurrentIndex(0)
        row_top_1.addWidget(self.cmb_launch_mode)
        row_top_1.addStretch()

        row_top_2 = QHBoxLayout()
        row_top_2.setSpacing(8)
        params_layout.addLayout(row_top_2)
        self.lbl_a_cars = QLabel("A数量：")
        row_top_2.addWidget(self.lbl_a_cars)
        self.spn_a_cars = QSpinBox()
        self.spn_a_cars.setRange(0, 9999)
        self.spn_a_cars.setValue(4)
        row_top_2.addWidget(self.spn_a_cars)

        row_top_2.addSpacing(10)
        self.lbl_b_cars = QLabel("B数量：")
        row_top_2.addWidget(self.lbl_b_cars)
        self.spn_b_cars = QSpinBox()
        self.spn_b_cars.setRange(0, 9999)
        self.spn_b_cars.setValue(0)
        row_top_2.addWidget(self.spn_b_cars)

        row_top_2.addSpacing(10)
        self.lbl_c_cars = QLabel("C数量：")
        row_top_2.addWidget(self.lbl_c_cars)
        self.spn_c_cars = QSpinBox()
        self.spn_c_cars.setRange(0, 9999)
        self.spn_c_cars.setValue(0)
        row_top_2.addWidget(self.spn_c_cars)
        row_top_2.addStretch()

        row_ratio = QHBoxLayout()
        row_ratio.setSpacing(8)
        params_layout.addLayout(row_ratio)
        self.lbl_total_cars = QLabel("分析时间：")
        row_ratio.addWidget(self.lbl_total_cars)
        self.spn_total_cars = QSpinBox()
        self.spn_total_cars.setRange(1, 9999)
        self.spn_total_cars.setValue(60)
        self.spn_total_cars.setSuffix(" 分钟")
        self.spn_total_cars.setToolTip(
            "按比例投车模式下使用；空线起步，第1台从0s进入首岗位，按分析时间和目标节拍计算理论投车台数。"
        )
        row_ratio.addWidget(self.spn_total_cars)
        self.lbl_total_cars.hide()
        self.spn_total_cars.hide()

        # 旧比例文本框保留但隐藏，后续可删除；当前按比例模式改用 A/B/C 数值框作为比例。
        self.ed_ratio = QLineEdit()
        self.ed_ratio.hide()

        row_ratio.addStretch()

        row_top_3 = QHBoxLayout()
        row_top_3.setSpacing(8)
        params_layout.addLayout(row_top_3)
        row_top_3.addWidget(QLabel("时间格刻度："))
        self.cmb_grid = QComboBox()
        self.cmb_grid.addItems(["1.0", "0.5", "2.0"])
        self.cmb_grid.setCurrentIndex(0)
        row_top_3.addWidget(self.cmb_grid)

        # 等待分配先隐藏，底层默认按“开始前等待”
        self.cmb_wait = QComboBox()
        self.cmb_wait.addItems(["开始前等待", "末尾等待"])
        self.cmb_wait.setCurrentIndex(0)
        self.cmb_wait.hide()

        row_top_3.addSpacing(12)
        row_top_3.addWidget(QLabel("目标节拍："))
        self.spn_target_takt = QSpinBox()
        self.spn_target_takt.setRange(0, 9999)
        self.spn_target_takt.setValue(118)
        self.spn_target_takt.setToolTip("0表示不进行节拍判定；大于0时按各岗位A/B/C实际工时判断OK/NG")
        row_top_3.addWidget(self.spn_target_takt)
        
        row_top_3.addSpacing(12)
        self.lbl_sequence_mode = QLabel("排列方式：")
        row_top_3.addWidget(self.lbl_sequence_mode)
        self.cmb_seq = QComboBox()
        self.cmb_seq.addItems(["顺排(A→B→C)", "交替混流"])
        self.cmb_seq.setCurrentIndex(0)
        row_top_3.addWidget(self.cmb_seq)

        row_top_3.addSpacing(12)
        self.lbl_max_run = QLabel("最大连续台数：")
        row_top_3.addWidget(self.lbl_max_run)
        self.spn_max_run = QSpinBox()
        self.spn_max_run.setRange(1, 9999)
        self.spn_max_run.setValue(10)
        self.spn_max_run.setToolTip("默认10台；填1表示尽量强制交替")
        row_top_3.addWidget(self.spn_max_run)
        row_top_3.addStretch()

        self.params_tip = QLabel(
            "顺排/交替混流：A/B/C 填数量；按比例投车：A/B/C 填比例，并填写分析时间。下方岗位矩阵用于逐行录入步骤。"
        )
        self.params_tip.setWordWrap(True)
        self.params_tip.setStyleSheet("color: #5f6b7a; font-size: 12px;")
        params_layout.addWidget(self.params_tip)

        # 中部：岗位矩阵区块
        table_frame, table_layout = _make_block("岗位矩阵")

        table_action_row = QHBoxLayout()
        table_action_row.setSpacing(8)
        table_layout.addLayout(table_action_row)
        table_action_row.addStretch()
        self.btn_add_row = QPushButton("添加步骤", self.page_multi_input)
        self.btn_del_row = QPushButton("删除步骤", self.page_multi_input)
        self.btn_fill_sample = QPushButton("填入示例", self.page_multi_input)
        table_action_row.addWidget(self.btn_add_row)
        table_action_row.addWidget(self.btn_del_row)
        table_action_row.addWidget(self.btn_fill_sample)

        input_layout.addWidget(table_frame, 1)

        input_next_row = QHBoxLayout()
        input_next_row.setSpacing(8)
        input_next_row.addStretch()
        self.btn_export = QPushButton("生成并导出组合票", self.page_multi_input)
        input_next_row.addWidget(self.btn_export)
        self.btn_go_result_page = QPushButton("下一步：分析与导出", self.page_multi_input)
        input_next_row.addWidget(self.btn_go_result_page)
        input_layout.addLayout(input_next_row)

        self.tbl = QTableWidget(0, 8, self)
        self.tbl.setHorizontalHeaderLabels([
            "序号", "工程名称", "设备数量", "所属线别",
            "岗位设备", "A工时", "B工时", "C工时"
        ])
        self.tbl.horizontalHeader().setStretchLastSection(False)
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
        )
        self.tbl.setTabKeyNavigation(False)
        self.tbl.setAlternatingRowColors(True)
        table_layout.addWidget(self.tbl, 1)

        # 分析页顶部：只保留一个小按钮行，避免占用纵向空间。
        result_nav_row = QHBoxLayout()
        result_nav_row.setSpacing(8)
        result_nav_row.addStretch()
        self.btn_analyze = QPushButton("分析当前排程", self.page_multi_result)
        result_nav_row.addWidget(self.btn_analyze)
        result_layout.addLayout(result_nav_row)

        # 以下控件仅作为旧逻辑兼容容器保留，不再占用分析页版面空间。
        self.lbl_analysis = QLabel(
            "结果分析：点击『分析当前排程』后显示总车数、总完成时间、总等待时间、平均等待时间与节拍判定。",
            self.page_multi_result,
        )
        self.lbl_analysis.setWordWrap(True)
        self.lbl_analysis.hide()

        self.tbl_station_analysis = QTableWidget(0, 8, self.page_multi_result)
        self.tbl_station_analysis.setHorizontalHeaderLabels([
            "岗位", "经过台数", "累计工时", "累计等待", "平均工时", "平均等待", "节拍判定", "超节拍车型"
        ])
        self.tbl_station_analysis.horizontalHeader().setStretchLastSection(False)
        self.tbl_station_analysis.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_station_analysis.verticalHeader().setVisible(False)
        self.tbl_station_analysis.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_station_analysis.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_station_analysis.setAlternatingRowColors(True)
        self.tbl_station_analysis.hide()

        # 车型数据摘要区：作为后续完整动画仿真的前置展示层
        vehicle_summary_frame, vehicle_summary_layout = _make_block("车型数据摘要区")
        result_layout.addWidget(vehicle_summary_frame, 1)

        sim_control_row = QHBoxLayout()
        sim_control_row.setSpacing(8)
        vehicle_summary_layout.addLayout(sim_control_row)
        self.lbl_sim_time = QLabel("仿真时间：0.0s / 0.0s")
        sim_control_row.addWidget(self.lbl_sim_time)
        sim_control_row.addSpacing(16)
        self.lbl_sim_total_wait = QLabel("总等待：0s")
        sim_control_row.addWidget(self.lbl_sim_total_wait)
        sim_control_row.addSpacing(16)
        self.sim_progress = QProgressBar(self.page_multi_result)
        self.sim_progress.setRange(0, 1000)
        self.sim_progress.setValue(0)
        self.sim_progress.setTextVisible(True)
        self.sim_progress.setFixedWidth(220)
        self.sim_progress.setFormat("进度 %p%")
        sim_control_row.addWidget(self.sim_progress)
        sim_control_row.addStretch()
        self.cmb_sim_speed = QComboBox()
        self.cmb_sim_speed.addItems(["1x", "5x", "10x", "50x"])
        self.cmb_sim_speed.setCurrentText("10x")
        self.btn_sim_play = QPushButton("播放")
        self.btn_sim_pause = QPushButton("暂停")
        self.btn_sim_reset = QPushButton("重置")
        sim_control_row.addWidget(QLabel("速度："))
        sim_control_row.addWidget(self.cmb_sim_speed)
        sim_control_row.addWidget(self.btn_sim_play)
        sim_control_row.addWidget(self.btn_sim_pause)
        sim_control_row.addWidget(self.btn_sim_reset)

        self.sim_timer = QTimer(self)
        self.sim_timer.setInterval(100)
        self.sim_time = 0.0
        self.last_schedule_rows = []
        self.last_analysis = None
        self.last_max_finish = 0.0

        self.lbl_vehicle_summary = QLabel("")
        self.lbl_vehicle_summary.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.lbl_vehicle_summary.setWordWrap(True)
        self.lbl_vehicle_summary.setTextFormat(Qt.RichText)
        self.lbl_vehicle_summary.setMaximumHeight(115)
        self.lbl_vehicle_summary.setStyleSheet(
            "background: #f7f9fc;"
            "border: 1px dashed #cbd5e1;"
            "border-radius: 8px;"
            "padding: 16px;"
            "color: #334155;"
            "font-size: 13px;"
            "line-height: 1.5;"
        )

        vehicle_summary_layout.addWidget(self.lbl_vehicle_summary, 1)

        self.sim_scene = QGraphicsScene(self)
        self.sim_graphics_view = QGraphicsView(self.sim_scene, self.page_multi_result)
        self.sim_graphics_view.setMinimumHeight(260)
        self.sim_graphics_view.setStyleSheet(
            "background: #f1f5f9;"
            "border: 1px solid #cbd5e1;"
            "border-radius: 10px;"
        )
        vehicle_summary_layout.addWidget(self.sim_graphics_view, 2)
        self.lbl_sim_view = QLabel("仿真画面：请先点击『分析当前排程』。")
        self.lbl_sim_view.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.lbl_sim_view.setWordWrap(True)
        self.lbl_sim_view.setMinimumHeight(90)
        self.lbl_sim_view.setMaximumHeight(130)
        self.lbl_sim_view.setStyleSheet(
            "background: #ffffff;"
            "border: 1px solid #e2e8f0;"
            "border-radius: 8px;"
            "padding: 12px;"
            "color: #334155;"
            "font-size: 12px;"
        )
        vehicle_summary_layout.addWidget(self.lbl_sim_view, 1)

        self.txt_schedule_debug = QPlainTextEdit(self.page_multi_result)
        self.txt_schedule_debug.setReadOnly(True)
        self.txt_schedule_debug.setMaximumBlockCount(300)
        self.txt_schedule_debug.setMaximumHeight(160)
        self.txt_schedule_debug.setPlaceholderText("排程运行日志：点击『分析当前排程』后显示前 200 条 rows 明细。")
        self.txt_schedule_debug.setStyleSheet(
            "background: #0f172a;"
            "border: 1px solid #334155;"
            "border-radius: 8px;"
            "padding: 8px;"
            "color: #e2e8f0;"
            "font-family: Menlo, Consolas, monospace;"
            "font-size: 11px;"
        )
        vehicle_summary_layout.addWidget(self.txt_schedule_debug, 1)

        # ---------- Tab2：单工程组合票 ----------
        self.page_single = QWidget(self)
        page_single_layout = QVBoxLayout(self.page_single)
        page_single_layout.setContentsMargins(8, 8, 8, 8)
        page_single_layout.setSpacing(8)

        # 顶部基本信息
        row_info = QHBoxLayout()
        page_single_layout.addLayout(row_info)

        row_info.addWidget(QLabel("工程名称："))
        self.ed_sw_project = QLineEdit(self.page_single)
        self.ed_sw_project.setPlaceholderText("例如：前轴调整工位")
        self.ed_sw_project.setFixedWidth(200)
        row_info.addWidget(self.ed_sw_project)

        row_info.addSpacing(12)
        row_info.addWidget(QLabel("品番·品名："))
        self.ed_sw_part = QLineEdit(self.page_single)
        self.ed_sw_part.setPlaceholderText("例如：XXXX-XXXXX 前轮定位")
        self.ed_sw_part.setFixedWidth(220)
        row_info.addWidget(self.ed_sw_part)

        row_info.addSpacing(12)
        row_info.addWidget(QLabel("作业者："))
        self.ed_sw_worker = QLineEdit(self.page_single)
        self.ed_sw_worker.setPlaceholderText("例如：张三")
        self.ed_sw_worker.setFixedWidth(120)
        row_info.addWidget(self.ed_sw_worker)

        row_info.addSpacing(12)
        row_info.addWidget(QLabel("节拍TT(秒)："))
        self.spn_sw_takt = QSpinBox(self.page_single)
        self.spn_sw_takt.setRange(1, 9999)
        self.spn_sw_takt.setValue(118)  # 默认示例
        row_info.addWidget(self.spn_sw_takt)

        row_info.addStretch()

        # 作业手顺表（A→B 区间）
        self.tbl_sw = QTableWidget(0, 8, self.page_single)
        self.tbl_sw.setHorizontalHeaderLabels([
            "顺序", "作业名称A", "作业名称B",
            "手作业(秒)", "自动(秒)", "步行(秒)",
            "步行在前/后", "自动在前/后"
        ])
        self.tbl_sw.horizontalHeader().setStretchLastSection(True)
        self.tbl_sw.verticalHeader().setVisible(False)
        page_single_layout.addWidget(self.tbl_sw, 1)

        # 底部按钮栏（单工程组合票）
        row_btn_sw = QHBoxLayout()
        page_single_layout.addLayout(row_btn_sw)
        row_btn_sw.addStretch()

        self.btn_sw_add = QPushButton("添加作业行", self.page_single)
        self.btn_sw_del = QPushButton("删除选中行", self.page_single)
        self.btn_sw_export = QPushButton("导出标准作业组合票", self.page_single)

        row_btn_sw.addWidget(self.btn_sw_add)
        row_btn_sw.addWidget(self.btn_sw_del)
        row_btn_sw.addWidget(self.btn_sw_export)

        self.tabs.addTab(self.page_single, "单工程组合票")

        # 状态栏（用于显示导出进度 / 完成信息）
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self._on_tab_changed(self.tabs.currentIndex())
        self._update_mode_ui()

    def _connect_signals(self):
        self.btn_add_row.clicked.connect(self.add_row)
        self.btn_del_row.clicked.connect(self.del_row)
        self.btn_fill_sample.clicked.connect(self.fill_sample)
        self.cmb_launch_mode.currentIndexChanged.connect(self._update_mode_ui)
        self.btn_go_result_page.clicked.connect(lambda: self.multi_tabs.setCurrentWidget(self.page_multi_result))
        self.btn_analyze.clicked.connect(self.do_analyze)
        self.btn_export.clicked.connect(self.do_export)
        self.act_help.triggered.connect(self.show_help)
        self.btn_sim_play.clicked.connect(self._start_simulation)
        self.btn_sim_pause.clicked.connect(self._pause_simulation)
        self.btn_sim_reset.clicked.connect(self._reset_simulation)
        self.sim_timer.timeout.connect(self._on_simulation_tick)

        # Tab 切换时，控制多工程页内步骤按钮是否可用
        self.tabs.currentChanged.connect(self._on_tab_changed)

        # 单工程组合票 Tab
        self.btn_sw_add.clicked.connect(self.add_single_row)
        self.btn_sw_del.clicked.connect(self.del_single_row)
        self.btn_sw_export.clicked.connect(self.export_single_placeholder)

    def _update_mode_ui(self):
        """根据模式切换 A/B/C 输入含义：数量模式填写数量，比例模式填写比例 + 分析时间。"""
        is_ratio = self.cmb_launch_mode.currentIndex() == 1
        if hasattr(self, "lbl_a_cars"):
            self.lbl_a_cars.setText("A比例：" if is_ratio else "A数量：")
        if hasattr(self, "lbl_b_cars"):
            self.lbl_b_cars.setText("B比例：" if is_ratio else "B数量：")
        if hasattr(self, "lbl_c_cars"):
            self.lbl_c_cars.setText("C比例：" if is_ratio else "C数量：")
        if hasattr(self, "lbl_total_cars"):
            self.lbl_total_cars.setVisible(is_ratio)
        if hasattr(self, "spn_total_cars"):
            self.spn_total_cars.setVisible(is_ratio)
        if hasattr(self, "lbl_sequence_mode"):
            self.lbl_sequence_mode.setVisible(not is_ratio)
        if hasattr(self, "cmb_seq"):
            self.cmb_seq.setVisible(not is_ratio)
        if hasattr(self, "lbl_max_run"):
            self.lbl_max_run.setVisible(not is_ratio)
        if hasattr(self, "spn_max_run"):
            self.spn_max_run.setVisible(not is_ratio)
            self.spn_max_run.setEnabled(not is_ratio)
        if hasattr(self, "params_tip"):
            if is_ratio:
                self.params_tip.setText(
                    "当前模式：按比例投车。A/B/C 填比例；分析时间填写 xx 分钟。"
                    "程序按空线起步模型，第1台从0s进入首岗位，并按目标节拍计算理论投车台数。"
                )
            else:
                self.params_tip.setText("当前模式：按数量投车。A/B/C 填数量；顺排按 A→B→C，交替混流可配合最大连续台数使用。")
    def do_analyze(self):
        try:
            project, cars, grid_step, wait_policy, defs, vehicle_counts, sequence_mode, max_consecutive, ratio_pattern, target_takt = self._collect_inputs()
            rows, max_finish = tickets.schedule(
                defs,
                cars,
                vehicle_counts,
                sequence_mode,
                max_consecutive,
                ratio_pattern,
                launch_takt=target_takt,
            )
            analysis = tickets.analyze_schedule(rows, max_finish, target_takt)
            analysis = self._apply_time_window_analysis(analysis, rows, target_takt)
            self.last_schedule_rows = rows
            self.last_analysis = analysis
            self.last_max_finish = float(max_finish or 0.0)
            if hasattr(self, "txt_schedule_debug"):
                self.txt_schedule_debug.setPlainText(self._build_schedule_debug_log(rows, limit=200))
            self.sim_time = 0.0
            self._update_sim_time_label()
            self._update_sim_total_wait_label()
            self._update_sim_view()
            self._draw_sim_scene()
            self._show_analysis_result(analysis)
            self.status.showMessage("排程分析完成", 6000)
        except Exception as e:
            import traceback
            detail = traceback.format_exc()
            print(detail)
            QMessageBox.warning(self, "分析失败", str(e))
    def _update_sim_total_wait_label(self):
        """刷新仿真控制栏中的关键判定信息。"""
        if not hasattr(self, "lbl_sim_total_wait"):
            return
        analysis = getattr(self, "last_analysis", None)
        summary = analysis.get("summary", {}) if isinstance(analysis, dict) else {}
        blocking_time = self._fmt_analysis_num(summary.get("total_wait", 0.0))
        self.lbl_sim_total_wait.setText(f"累计阻塞：{blocking_time}s")
    def _sim_speed_value(self) -> float:
        """读取仿真倍速。"""
        if not hasattr(self, "cmb_sim_speed"):
            return 1.0
        text = self.cmb_sim_speed.currentText().replace("x", "").strip()
        try:
            return max(1.0, float(text))
        except Exception:
            return 1.0

    def _update_sim_time_label(self):
        """刷新仿真时间显示。"""
        if not hasattr(self, "lbl_sim_time"):
            return
        current = float(getattr(self, "sim_time", 0.0) or 0.0)
        total = float(getattr(self, "last_max_finish", 0.0) or 0.0)
        self.lbl_sim_time.setText(f"仿真时间：{current:.1f}s / {total:.1f}s")
        if hasattr(self, "sim_progress"):
            if total > 0:
                progress_value = int(max(0.0, min(1.0, current / total)) * 1000)
            else:
                progress_value = 0
            self.sim_progress.setValue(progress_value)

    def _start_simulation(self):
        """启动仿真计时。当前阶段只推进时间，不绘制车辆。"""
        if float(getattr(self, "last_max_finish", 0.0) or 0.0) <= 0:
            QMessageBox.information(self, "提示", "请先点击『分析当前排程』生成仿真数据。")
            return
        self.sim_timer.start()

    def _pause_simulation(self):
        """暂停仿真计时。"""
        if hasattr(self, "sim_timer"):
            self.sim_timer.stop()

    def _reset_simulation(self):
        """重置仿真时间。"""
        if hasattr(self, "sim_timer"):
            self.sim_timer.stop()
        self.sim_time = 0.0
        self._update_sim_time_label()
        self._update_sim_view()
        self._draw_sim_scene()

    def _on_simulation_tick(self):
        """仿真计时推进。后续车辆绘制会基于 sim_time 刷新画面。"""
        total = float(getattr(self, "last_max_finish", 0.0) or 0.0)
        if total <= 0:
            self._pause_simulation()
            return
        self.sim_time = min(total, float(getattr(self, "sim_time", 0.0) or 0.0) + 0.1 * self._sim_speed_value())
        self._update_sim_time_label()
        self._update_sim_view()
        self._draw_sim_scene()
        if self.sim_time >= total:
            self._pause_simulation()


    def _sim_row_value(self, row: dict, *keys, default=None):
        """从排程行中兼容读取字段。"""
        for key in keys:
            if key in row and row.get(key) is not None:
                return row.get(key)
        return default

    def _sim_row_start(self, row: dict) -> float:
        value = self._sim_row_value(row, "start", "start_time", "begin", "in", "in_time", default=0.0)
        try:
            return float(value or 0.0)
        except Exception:
            return 0.0

    def _sim_row_end(self, row: dict) -> float:
        value = self._sim_row_value(row, "end", "finish", "finish_time", "out", "out_time", default=None)
        if value is not None:
            try:
                return float(value or 0.0)
            except Exception:
                pass
        start = self._sim_row_start(row)
        dur = self._sim_row_value(row, "dur", "duration", "process", "process_time", default=0.0)
        try:
            return start + float(dur or 0.0)
        except Exception:
            return start

    def _sim_row_station(self, row: dict) -> str:
        return str(self._sim_row_value(row, "step_display", "station", "display", "name", "group", default="岗位") or "岗位")

    def _sim_row_car_label(self, row: dict) -> str:
        car = self._sim_row_value(row, "car", "car_no", "car_index", "idx", default="?")
        car_type = str(self._sim_row_value(row, "car_type", "type", "vehicle_type", default="") or "")
        if car_type:
            return f"Car#{car}({car_type})"
        return f"Car#{car}"
    
    def _sim_row_run_mode(self, row: dict) -> str:
        """兼容读取岗位运行方式。"""
        return str(self._sim_row_value(row, "run_mode", "mode", default="") or "")

    def _sim_row_line_no(self, row: dict) -> str:
        """兼容读取排程行线别，用于 v2-4 线别验证。"""
        return str(self._sim_row_value(row, "line_no", "line", "line_scope", default="") or "")

    def _build_schedule_debug_log(self, rows, limit: int = 200) -> str:
        """临时排程运行日志，用于 v2 排程模型验证；后续稳定后可隐藏或删除。"""
        if not rows:
            return "暂无排程 rows。"

        lines = [
            "排程运行日志（临时验证用）",
            "字段：Car | 岗位 | 线别 | 资源 | 理论投车 | 开始 | 完成 | 离开 | 工时 | 等待",
            "-" * 120,
        ]

        def _fmt(value):
            try:
                return f"{float(value or 0.0):.1f}"
            except Exception:
                return str(value)

        sorted_rows = sorted(
            list(rows or []),
            key=lambda r: (
                int(r.get("car", r.get("car_no", r.get("car_index", 0))) or 0),
                float(r.get("start", r.get("start_time", 0.0)) or 0.0),
                int(r.get("step_seq", 0) or 0),
            ),
        )

        for idx, row in enumerate(sorted_rows[:limit], start=1):
            car = row.get("car", row.get("car_no", row.get("car_index", "?")))
            car_type = str(row.get("car_type", row.get("duration_source", row.get("vehicle_type", ""))) or "")
            station = str(row.get("step_display", row.get("station", row.get("group", "岗位"))) or "岗位")
            line_no = str(row.get("line_no", row.get("line", "")) or "")
            resource_key = str(row.get("resource_key", "") or "")
            theory_launch = row.get("theory_launch_time", "")
            start = row.get("start", row.get("start_time", 0.0))
            svc_finish = row.get("svc_finish", row.get("finish", row.get("end", 0.0)))
            depart = row.get("depart", row.get("end", row.get("svc_finish", 0.0)))
            dur = row.get("dur", row.get("duration", 0.0))
            wait = float(row.get("block_wait", 0.0) or 0.0) + float(row.get("launch_wait", 0.0) or 0.0)

            lines.append(
                f"{idx:03d}. Car#{car}({car_type}) | {station} | "
                f"线别 {line_no or '—'} | 资源 {resource_key or '—'} | "
                f"理论投车 {_fmt(theory_launch)}s | 开始 {_fmt(start)}s | "
                f"完成 {_fmt(svc_finish)}s | 离开 {_fmt(depart)}s | "
                f"工时 {_fmt(dur)}s | 等待 {_fmt(wait)}s"
            )

        if len(sorted_rows) > limit:
            lines.append(f"……仅显示前 {limit} 条，共 {len(sorted_rows)} 条。")

        return "\n".join(lines)
    
    def _sim_car_key(self, row: dict):
        """按车辆编号聚合排程段。"""
        return self._sim_row_value(row, "car", "car_no", "car_index", "idx", default="?")
    
    def _sim_car_rows(self):
        """将排程行按车辆聚合，并按开始时间排序。"""
        grouped = {}
        for row in getattr(self, "last_schedule_rows", []) or []:
            key = self._sim_car_key(row)
            grouped.setdefault(key, []).append(row)
        for key in grouped:
            grouped[key].sort(key=lambda r: self._sim_row_start(r))
        return grouped

    def _sim_station_names(self):
        """按排程行出现顺序提取岗位名。"""
        names = []
        seen = set()
        for row in getattr(self, "last_schedule_rows", []) or []:
            name = self._sim_row_station(row)
            if name not in seen:
                names.append(name)
                seen.add(name)
        return names

    def _update_sim_view(self):
        """刷新简易仿真画面。当前阶段用文字展示岗位轨道与加工中的车辆。"""
        if not hasattr(self, "lbl_sim_view"):
            return

        rows = getattr(self, "last_schedule_rows", []) or []
        if not rows:
            self.lbl_sim_view.setText("仿真画面：请先点击『分析当前排程』。")
            return

        current = float(getattr(self, "sim_time", 0.0) or 0.0)
        station_names = self._sim_station_names()
        station_line = " → ".join([f"ST-{i + 1} {name}" for i, name in enumerate(station_names)])

        active_rows = []
        next_rows = []
        for row in rows:
            start = self._sim_row_start(row)
            end = self._sim_row_end(row)
            if start <= current < end:
                active_rows.append((start, end, row))
            elif start > current:
                next_rows.append((start, end, row))

        active_rows.sort(key=lambda x: (self._sim_row_station(x[2]), x[0]))
        next_rows.sort(key=lambda x: x[0])

        lines = [
            f"当前时间：{current:.1f}s",
            f"岗位轨道：{station_line or '—'}",
            "",
            "加工中车辆：",
        ]

        if active_rows:
            for start, end, row in active_rows[:8]:
                car_label = self._sim_row_car_label(row)
                station = self._sim_row_station(row)
                dur = max(0.0, end - start)
                remain = max(0.0, end - current)
                line_no = self._sim_row_line_no(row)
                line_suffix = f" ｜ 线别 {line_no}" if line_no else ""
                run_mode = self._sim_row_run_mode(row)
                mode_suffix = f" ｜ {run_mode}" if run_mode else ""
                lines.append(f"- PROC {car_label} @ {station}{line_suffix}{mode_suffix} ｜ 开始 {start:.1f}s ｜ 加工 {dur:.1f}s ｜ 剩余 {remain:.1f}s")
            if len(active_rows) > 8:
                lines.append(f"- ……还有 {len(active_rows) - 8} 台/段正在加工")
        else:
            lines.append("- 当前无车辆处于加工中")

        lines.append("")
        lines.append("下一段即将开始：")
        if next_rows:
            for start, end, row in next_rows[:3]:
                car_label = self._sim_row_car_label(row)
                station = self._sim_row_station(row)
                line_no = self._sim_row_line_no(row)
                line_suffix = f" ｜ 线别 {line_no}" if line_no else ""
                lines.append(f"- {car_label} @ {station}{line_suffix} ｜ {start:.1f}s 开始")
        else:
            lines.append("- 已无后续加工段")

        self.lbl_sim_view.setText("\n".join(lines))

    # ------------- 多车组合票：动作 ------------- #
    def add_row(self):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)

        # 序号
        self.tbl.setItem(r, 0, QTableWidgetItem(str(r + 1)))
        # 工程名称
        self.tbl.setItem(r, 1, QTableWidgetItem(""))

        # 设备数量：v2 主线字段。默认 2，代表双线双设备。
        device_count_cb = QComboBox(self.tbl)
        device_count_cb.addItems(["1", "2"])
        device_count_cb.setCurrentText("2")
        self.tbl.setCellWidget(r, 2, device_count_cb)

        # 所属线别：v2 主线字段。
        line_scope_cb = QComboBox(self.tbl)
        line_scope_cb.addItems(["1号线", "2号线", "双线", "双线共用"])
        line_scope_cb.setCurrentText("双线")
        self.tbl.setCellWidget(r, 3, line_scope_cb)

        # 岗位设备
        self.tbl.setItem(r, 4, QTableWidgetItem(""))
        # A / B / C 工时
        self.tbl.setItem(r, 5, QTableWidgetItem(""))
        self.tbl.setItem(r, 6, QTableWidgetItem(""))
        self.tbl.setItem(r, 7, QTableWidgetItem(""))

        def _sync_line_scope():
            if device_count_cb.currentText() == "2":
                line_scope_cb.setCurrentText("双线")
                line_scope_cb.setEnabled(False)
            else:
                line_scope_cb.setEnabled(True)
                if line_scope_cb.currentText() == "双线":
                    line_scope_cb.setCurrentText("1号线")

        device_count_cb.currentTextChanged.connect(_sync_line_scope)
        _sync_line_scope()

    def _choose_color(self, row: int):
        dlg_col = QColorDialog.getColor(parent=self)
        if dlg_col.isValid():
            hex_code = dlg_col.name()
            btn = self.tbl.cellWidget(row, self.COL_C_TIME)
            btn.setStyleSheet(f"background:{hex_code};")
            self.tbl.item(row, self.COL_C_TIME).setData(Qt.UserRole, hex_code)

    def del_row(self):
        r = self.tbl.currentRow()
        if r >= 0:
            self.tbl.removeRow(r)

    # -------- 单人标准作业组合票：行操作 --------
    def add_single_row(self):
        """在单人作业手顺表中新增一行"""
        if not hasattr(self, "tbl_sw"):
            return

        current_rows = self.tbl_sw.rowCount()
        max_steps = getattr(self, "MAX_SINGLE_STEPS", 23)

        if current_rows >= max_steps:
            QMessageBox.warning(
                self,
                "已到模板上限",
                f"当前单人标准作业组合票模板最多支持 {max_steps} 行。\n"
                f"你现在已经添加了 {current_rows} 行，不能再继续新增。\n\n"
                "请合并部分区间或拆分为多张组合票后再导出。"
            )
            return

        r = current_rows
        self.tbl_sw.insertRow(r)
        # 顺序默认递增（组合票行号）
        self.tbl_sw.setItem(r, 0, QTableWidgetItem(str(r + 1)))
        # 作业名称A / B 先留空，让你填写
        self.tbl_sw.setItem(r, 1, QTableWidgetItem(""))
        self.tbl_sw.setItem(r, 2, QTableWidgetItem(""))
        # 手作业 / 自动 / 步行，默认 0
        self.tbl_sw.setItem(r, 3, QTableWidgetItem("0"))
        self.tbl_sw.setItem(r, 4, QTableWidgetItem("0"))
        self.tbl_sw.setItem(r, 5, QTableWidgetItem("0"))
        # 步行位置：默认“后置”
        pos_cb = QComboBox(self.tbl_sw)
        pos_cb.addItem("后置", userData="after")
        pos_cb.addItem("前置", userData="before")
        self.tbl_sw.setCellWidget(r, 6, pos_cb)
        # 自动在前/后（默认后置）
        auto_cb = QComboBox(self.tbl_sw)
        auto_cb.addItem("后置", userData="after")
        auto_cb.addItem("前置", userData="before")
        self.tbl_sw.setCellWidget(r, 7, auto_cb)

    def del_single_row(self):
        """删除单人作业手顺表中的选中行"""
        if not hasattr(self, "tbl_sw"):
            return
        r = self.tbl_sw.currentRow()
        if r >= 0:
            self.tbl_sw.removeRow(r)
        # 重写顺序列，让它保持 1,2,3,...
        for i in range(self.tbl_sw.rowCount()):
            item = self.tbl_sw.item(i, 0)
            if item is None:
                item = QTableWidgetItem()
                self.tbl_sw.setItem(i, 0, item)
            item.setText(str(i + 1))

    # -------- 单人标准作业组合票：数据收集 --------
    def _collect_single_inputs(self):
        """
        从单人作业手顺 Tab 中读取数据，并计算时间汇总。
        返回：
          project, part, worker, takt_sec, steps, totals
        其中：
          steps: [{seq, name, name_a, name_b, manual, auto, walk, walk_pos, auto_pos, duration, start, end}, ...]
          totals: {"manual": x, "auto": y, "walk": z, "total": t}
        """
        if not hasattr(self, "tbl_sw"):
            raise ValueError("单人作业手顺表尚未初始化")

        project = (self.ed_sw_project.text().strip() or "工程")
        part = self.ed_sw_part.text().strip()
        worker = self.ed_sw_worker.text().strip()
        takt_sec = int(self.spn_sw_takt.value())

        steps = []
        cur_time = 0.0
        total_manual = 0.0
        total_auto = 0.0
        total_walk = 0.0

        for r in range(self.tbl_sw.rowCount()):
            # 作业名称 A / B
            name_a_item = self.tbl_sw.item(r, 1)
            name_b_item = self.tbl_sw.item(r, 2)
            name_a = name_a_item.text().strip() if name_a_item else ""
            name_b = name_b_item.text().strip() if name_b_item else ""

            if not name_a and not name_b:
                # 两个都没填，当作空行，跳过
                continue

            # 导出时使用的显示名（A→B / 单独一个）
            if name_a and name_b:
                name = f"{name_a} → {name_b}"
            else:
                name = name_a or name_b

            def _get_time(col_idx: int) -> float:
                item = self.tbl_sw.item(r, col_idx)
                txt = item.text().strip() if item else ""
                if not txt:
                    return 0.0
                try:
                    return float(txt)
                except Exception:
                    raise ValueError(f"第 {r + 1} 行时间列（第 {col_idx + 1} 列）不是有效数字：{txt}")

            # 手作业 / 自动 / 步行时间列：3, 4, 5
            manual = _get_time(3)
            auto = _get_time(4)
            walk = _get_time(5)

            # 步行位置：前置/后置（默认后置）
            walk_pos = "after"
            pos_widget = self.tbl_sw.cellWidget(r, 6)
            if isinstance(pos_widget, QComboBox):
                walk_pos_data = pos_widget.currentData()
                if walk_pos_data in ("before", "after"):
                    walk_pos = walk_pos_data

            # 自动在前/后（默认后置）
            auto_pos = "after"
            auto_widget = self.tbl_sw.cellWidget(r, 7)
            if isinstance(auto_widget, QComboBox):
                auto_pos_data = auto_widget.currentData()
                if auto_pos_data in ("before", "after"):
                    auto_pos = auto_pos_data

            duration = manual + auto + walk
            if duration <= 0:
                raise ValueError(f"第 {r + 1} 行『{name}』的时间合计为 0，请填写手作业/自动/步行时间。")

            start = cur_time
            end = cur_time + duration
            cur_time = end

            total_manual += manual
            total_auto += auto
            total_walk += walk

            # 顺序列（如果用户改过，我们尽量读取）
            seq_item = self.tbl_sw.item(r, 0)
            try:
                seq = int(seq_item.text()) if seq_item and seq_item.text().strip() else len(steps) + 1
            except Exception:
                seq = len(steps) + 1

            steps.append({
                "seq": seq,
                "name": name,       # A→B 组合显示名（保留）
                "name_a": name_a,   # 原始作业名称A
                "name_b": name_b,   # 原始作业名称B
                "manual": manual,
                "auto": auto,
                "walk": walk,
                "walk_pos": walk_pos,  # 步行在前/后
                "auto_pos": auto_pos,  # 自动在前/后
                "duration": duration,
                "start": start,
                "end": end,
            })

        # 行数上限检查：防止超过模板预留的行数
        if len(steps) > self.MAX_SINGLE_STEPS:
            raise ValueError(
                f"当前单人标准作业组合票共有 {len(steps)} 行，已超过模板最多支持的 {self.MAX_SINGLE_STEPS} 行。\n"
                "请合并部分区间或拆分为多张组合票后再导出。"
            )

        if not steps:
            raise ValueError("请至少填写一行有效的作业（需有作业名称和时间）。")

        totals = {
            "manual": total_manual,
            "auto": total_auto,
            "walk": total_walk,
            "total": total_manual + total_auto + total_walk,
        }
        return project, part, worker, takt_sec, steps, totals

    # -------- 单人标准作业组合票：写入模板 --------
    def _export_single_to_excel(self, path, project, part, worker, takt_sec, steps, totals):
        """
        根据单人作业手顺（A→B 区间）将数据写入《组合票标准版.xlsx》模板：
        - 模板文件需放在与本文件同一目录下，文件名：组合票标准版.xlsx
        - 仅填充左侧步骤表区域和基本信息，不修改模板中的其他格式/图表
        """
        # 定位模板文件：与本 .py 同目录
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "组合票标准版.xlsx")
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"未找到模板文件：{template_path}")

        wb = load_workbook(template_path)
        try:
            ws = wb["④标准作业组合票"]
        except KeyError:
            ws = wb.active

        def _set_value(coord, value):
            """安全写入单元格：若目标是合并单元格，从其合并区域左上角写入"""
            cell = ws[coord]
            if isinstance(cell, MergedCell):
                for mr in ws.merged_cells.ranges:
                    if cell.coordinate in mr:
                        ws.cell(row=mr.min_row, column=mr.min_col).value = value
                        break
            else:
                cell.value = value

        def _set_fill(row, col, fill):
            """安全设置单元格填充：若目标是合并单元格，则写到其合并区域左上角"""
            cell = ws.cell(row=row, column=col)
            if isinstance(cell, MergedCell):
                for mr in ws.merged_cells.ranges:
                    if cell.coordinate in mr:
                        ws.cell(row=mr.min_row, column=mr.min_col).fill = fill
                        break
            else:
                cell.fill = fill

        def _set_border(row, col, border: Border):
            """
            安全设置单元格边框：若目标是合并单元格，则写到其合并区域左上角；
            与已有边框合并（只改指定方向的线型）。
            """
            cell = ws.cell(row=row, column=col)
            if isinstance(cell, MergedCell):
                for mr in ws.merged_cells.ranges:
                    if cell.coordinate in mr:
                        cell = ws.cell(row=mr.min_row, column=mr.min_col)
                        break

            old = cell.border or Border()

            def merge_side(new_side, old_side):
                if getattr(new_side, "style", None):
                    return new_side
                return old_side

            cell.border = Border(
                left=merge_side(border.left, old.left),
                right=merge_side(border.right, old.right),
                top=merge_side(border.top, old.top),
                bottom=merge_side(border.bottom, old.bottom),
                diagonal=old.diagonal,
                diagonal_direction=old.diagonal_direction,
                outline=old.outline,
                vertical=old.vertical,
                horizontal=old.horizontal,
            )

        def _clear_top_border(row, col):
            """
            清除单元格的上边框（保留其余边框），合并单元格时操作左上角单元格。
            """
            cell = ws.cell(row=row, column=col)
            # Handle merged cells: always operate on effective top-left cell
            if isinstance(cell, MergedCell):
                for mr in ws.merged_cells.ranges:
                    if cell.coordinate in mr:
                        cell = ws.cell(row=mr.min_row, column=mr.min_col)
                        break
            old = cell.border or Border()
            cell.border = Border(
                left=old.left,
                right=old.right,
                top=Side(style=None),
                bottom=old.bottom,
                diagonal=old.diagonal,
                diagonal_direction=old.diagonal_direction,
                outline=old.outline,
                vertical=old.vertical,
                horizontal=old.horizontal,
            )

        def _clear_left_border(row, col):
            """
            清除单元格的左边框（保留其余边框）；合并单元格时操作左上角单元格。
            """
            cell = ws.cell(row=row, column=col)
            # Handle merged cells: always operate on effective top-left cell
            if isinstance(cell, MergedCell):
                for mr in ws.merged_cells.ranges:
                    if cell.coordinate in mr:
                        cell = ws.cell(row=mr.min_row, column=mr.min_col)
                        break
            old = cell.border or Border()
            cell.border = Border(
                left=Side(style=None),
                right=old.right,
                top=old.top,
                bottom=old.bottom,
                diagonal=old.diagonal,
                diagonal_direction=old.diagonal_direction,
                outline=old.outline,
                vertical=old.vertical,
                horizontal=old.horizontal,
            )

        def _clear_right_border(row, col):
            """
            清除单元格的右边框（保留其余边框）；合并单元格时操作左上角单元格。
            """
            cell = ws.cell(row=row, column=col)
            # Handle merged cells: always operate on effective top-left cell
            if isinstance(cell, MergedCell):
                for mr in ws.merged_cells.ranges:
                    if cell.coordinate in mr:
                        cell = ws.cell(row=mr.min_row, column=mr.min_col)
                        break
            old = cell.border or Border()
            cell.border = Border(
                left=old.left,
                right=Side(style=None),
                top=old.top,
                bottom=old.bottom,
                diagonal=old.diagonal,
                diagonal_direction=old.diagonal_direction,
                outline=old.outline,
                vertical=old.vertical,
                horizontal=old.horizontal,
            )

        # 1) 清空左侧原有数据区域
        start_row = 9
        row_span = 3
        max_steps = getattr(self, "MAX_SINGLE_STEPS", 23)
        end_row = start_row + max_steps * row_span - 1
        for row in range(start_row, end_row + 1):
            for col in range(1, 6):
                cell = ws.cell(row=row, column=col)
                if isinstance(cell, MergedCell):
                    continue
                cell.value = None

        # 清空右侧时间轴区域填充（F列开始，按总时间估算范围）
        time_start_col = 6  # F列
        max_time = int(round(totals.get("total", 0))) if isinstance(totals, dict) else 0
        if max_time < 0:
            max_time = 0
        time_end_col = time_start_col + max_time + 5
        for row in range(start_row, end_row + 1):
            for col in range(time_start_col, time_end_col + 1):
                cell = ws.cell(row=row, column=col)
                if isinstance(cell, MergedCell):
                    continue
                cell.fill = PatternFill()

          # 2) 写入步骤：每步占 3 行（A9:A11, A12:A14, ...）
        row_span = 3
        time_start_col = 6  # F列，时间轴起始列
        time_fill = PatternFill(fill_type="solid", fgColor="000000")
        # 自动：加粗虚线（仅画上边框，不填充）
        auto_side = Side(style="mediumDashed", color="000000")
        h_auto_border = Border(top=auto_side)

        # segments：按“步骤”记录每一行的黑条起止（即手作业+自动的时间段）
        segments = []  # [{"mid_row": int, "bar_start": int, "bar_end": int}, ...]

        for idx, s in enumerate(steps):
            base_row = start_row + idx * row_span

            # 序号
            ws.cell(row=base_row, column=1).value = s["seq"]

            # 作业名称 A/B：B 列两行
            name_a = s.get("name_a") or s.get("name") or ""
            name_b = s.get("name_b") or ""
            ws.cell(row=base_row, column=2).value = name_a
            if name_b:
                ws.cell(row=base_row + 2, column=2).value = name_b

            # 时间数值（C~E）
            ws.cell(row=base_row, column=3).value = s["manual"]
            ws.cell(row=base_row, column=4).value = s["auto"]
            ws.cell(row=base_row, column=5).value = s["walk"]

            # ===== 时间轴绘制（手作业=黑填充；自动=加粗虚线；步行仅用折线表示） =====
            start_sec = int(round(s["start"]))
            manual = float(s["manual"])
            auto = float(s["auto"])
            walk = float(s["walk"])
            walk_pos = s.get("walk_pos", "after")
            auto_pos = s.get("auto_pos", "after")

            # 起点：若步行在前，整体右移
            if walk_pos == "before":
                bar_start_sec = int(round(start_sec + walk))
            else:
                bar_start_sec = start_sec

            mid_row = base_row + 1

            # 决定绘制顺序：自动在前/后
            draw_seq = []
            if auto_pos == "before":
                if auto > 0:
                    draw_seq.append(("auto", auto))
                if manual > 0:
                    draw_seq.append(("manual", manual))
            else:
                if manual > 0:
                    draw_seq.append(("manual", manual))
                if auto > 0:
                    draw_seq.append(("auto", auto))

            seg_start = bar_start_sec
            for kind, length in draw_seq:
                seg_end = int(round(seg_start + length))
                if seg_end > seg_start:
                    for sec in range(seg_start, seg_end):
                        col = time_start_col + sec
                        if kind == "manual":
                            _set_fill(mid_row, col, time_fill)        # 手作业：黑色填充
                        else:
                            _set_border(mid_row, col, h_auto_border)  # 自动：加粗虚线（上边框）
                seg_start = seg_end

            # 记录该步的整体开始/结束（不包含步行在后）
            bar_end_sec = seg_start
            segments.append(
                {
                    "mid_row": mid_row,
                    "bar_start": bar_start_sec,
                    "bar_end": bar_end_sec,
                }
            )

        # 2.5) 相邻「步骤」之间画连接线：
        #      - 有间隔：步行 → 实折线，从黑条末端右边一格开始，先竖后横
        #      - 无间隔：直接接续 → 加粗实直线
        if len(segments) >= 2:
            solid_side = Side(style="medium", color="000000")   # 加粗实线
            walk_side  = Side(style="medium", color="000000")   # 步行：加粗实线

            h_walk_border = Border(top=walk_side)               # 步行横线
            v_walk_left   = Border(left=walk_side)              # 竖线（当前列左边）
            v_walk_right  = Border(right=walk_side)             # 竖线镜像（前一列右边）
            v_solid_right_border = Border(right=solid_side)     # 无间隔直连竖线（边界线上）

            for i in range(len(segments) - 1):
                curr = segments[i]
                nxt = segments[i + 1]

                mid_row_curr = curr["mid_row"]
                mid_row_nxt = nxt["mid_row"]
                bar_end_curr = curr["bar_end"]
                bar_start_nxt = nxt["bar_start"]

                # 注意：bar_end / bar_start 是“时间（秒）”，还没加上 F 列偏移
                if bar_start_nxt > bar_end_curr:
                    # 有间隔：步行 → 实折线
                    # 连接策略：
                    #   - 竖线画在“上一列的右边界”，并且只画到下一段所在行的上一行（不进入下一段单元格）
                    #   - 横线从拐点所在列开始，沿下一段所在行的上边框一直画到下一段条形左侧
                    first_blank_col = time_start_col + bar_end_curr        # 上一段末尾右侧的第一格
                    next_bar_first_col = time_start_col + bar_start_nxt    # 下一段条形开始列

                    # 1) 竖线：用上一列（first_blank_col - 1）的『右边界』画，恰好停在下一段顶边
                    grid_col_for_right_edge = first_blank_col - 1
                    row_vert_start = mid_row_curr + 1
                    row_vert_end_exclusive = mid_row_nxt  # 不包含下一段所在行，避免出现“下垂尾巴”
                    if grid_col_for_right_edge >= time_start_col and row_vert_start < row_vert_end_exclusive:
                        for row in range(row_vert_start, row_vert_end_exclusive):
                            _set_border(row, grid_col_for_right_edge, v_walk_right)
                    # 保底清理下一行该列的右边界，避免‘下垂尾巴’
                    _clear_right_border(mid_row_nxt, grid_col_for_right_edge)

                    # 2) 横线：从拐点所在列开始（不跳空），一直到下一段左侧列
                    start_h_col = first_blank_col  # 不留缺口
                    if start_h_col < next_bar_first_col:
                        for col in range(start_h_col, next_bar_first_col):
                            _set_border(mid_row_nxt, col, h_walk_border)
                else:
                    # 无间隔：在上一段最后一秒所在列的“右边界”连线，
                    # 竖线落在列缝而不是下一段条形内部，且不覆盖下一段所在行
                    boundary_col = time_start_col + bar_end_curr
                    row_top = min(mid_row_curr, mid_row_nxt)
                    row_bottom = max(mid_row_curr, mid_row_nxt) - 1
                    if row_top <= row_bottom:
                        for row in range(row_top, row_bottom + 1):
                            _set_border(row, boundary_col - 1, v_solid_right_border)
                    # 保底清理下一行该列的右边界，避免‘下垂尾巴’
                    _clear_right_border(mid_row_nxt, boundary_col - 1)

        # 3) 合计行：B79 总时间，C79 手作业时间，D79 自动时间，E79 步行时间
        if isinstance(totals, dict):
            total_sec = totals.get("total", 0.0)
            manual_sec = totals.get("manual", 0.0)
            auto_sec = totals.get("auto", 0.0)
            walk_sec = totals.get("walk", 0.0)
        else:
            total_sec = manual_sec = auto_sec = walk_sec = 0.0

        def _fmt_sec(v):
            """把秒数统一转成整数秒写入单元格"""
            try:
                return int(round(float(v)))
            except Exception:
                return v

        _set_value("B79", _fmt_sec(total_sec))   # 合计下面：总时间
        _set_value("C79", _fmt_sec(manual_sec))  # 手作业合计
        _set_value("D79", _fmt_sec(auto_sec))    # 自动合计
        _set_value("E79", _fmt_sec(walk_sec))    # 步行合计

        # 4) 在上方空白处写入工程信息
        _set_value("B2", project)
        _set_value("B3", part)
        _set_value("B4", worker)
        _set_value("E2", takt_sec)

        # 5) 保存
        wb.save(path)

    def export_single_placeholder(self):
        """
        单工程组合票导出流程：
        1. 读取 Tab2 中 A→B 区间作业数据并校验
        2. 选择保存路径
        3. 使用固定 Excel 模板导出标准作业组合票
        """
        try:
            project, part, worker, takt_sec, steps, totals = self._collect_single_inputs()
        except Exception as e:
            QMessageBox.warning(self, "输入有误", str(e))
            return

        default_name = f"{project}_单人组合票.xlsx" if project else "单人组合票.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出标准作业组合票",
            default_name,
            "Excel (*.xlsx)",
        )
        if not path:
            return

        try:
            self._export_single_to_excel(path, project, part, worker, takt_sec, steps, totals)
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))
            return

        msg = (
            f"已导出标准作业组合票：\n{path}\n\n"
            f"工程名称：{project}\n"
            f"品番·品名：{part or '（未填写）'}\n"
            f"作业者：{worker or '（未填写）'}\n\n"
            f"节拍 TT：{takt_sec} 秒\n"
            f"总时间：{totals['total']:.1f} 秒\n"
            f"  其中 手作业：{totals['manual']:.1f} 秒\n"
            f"       自动：{totals['auto']:.1f} 秒\n"
            f"       步行：{totals['walk']:.1f} 秒\n\n"
            f"步骤数：{len(steps)} 步"
        )
        QMessageBox.information(self, "单工程组合票 - 导出完成", msg)

    # -------- 多车组合票：数据收集 & 导出 --------
    def fill_sample(self):
        """
        排程模型 v2 示例：设备数量 + 所属线别 + A/B/C 工时。
        工时 > 0 表示该车型经过该岗位；工时 = 0 表示该车型跳过该岗位。
        """
        self.tbl.setRowCount(0)
        sample_rows = [
            # 序号, 工程名称,   设备数量, 所属线别, 岗位设备,       A工时, B工时, C工时
            ("1",  "电集",       "2",    "双线",     "电集",         "100", "100", "100"),
            ("2",  "空悬+快充",  "1",    "1号线",    "空悬快充设备", "0",   "200", "0"),
            ("3",  "四轮定位",   "2",    "双线",     "四轮定位",     "110", "110", "110"),
        ]
        for row in sample_rows:
            self.add_row()
            r = self.tbl.rowCount() - 1

            self.tbl.setItem(r, 0, QTableWidgetItem(str(row[0])))
            self.tbl.setItem(r, 1, QTableWidgetItem(str(row[1])))

            device_count_widget = self.tbl.cellWidget(r, 2)
            if isinstance(device_count_widget, QComboBox):
                device_count_widget.setCurrentText(str(row[2]))

            line_scope_widget = self.tbl.cellWidget(r, 3)
            if isinstance(line_scope_widget, QComboBox):
                line_scope_widget.setCurrentText(str(row[3]))

            self.tbl.setItem(r, 4, QTableWidgetItem(str(row[4])))
            self.tbl.setItem(r, 5, QTableWidgetItem(str(row[5])))
            self.tbl.setItem(r, 6, QTableWidgetItem(str(row[6])))
            self.tbl.setItem(r, 7, QTableWidgetItem(str(row[7])))

        if not self.ed_project.text().strip():
            self.ed_project.setText("排程模型v2示例")
        self.spn_a_cars.setValue(4)
        self.spn_b_cars.setValue(2)
        self.spn_c_cars.setValue(0)
        self.spn_total_cars.setValue(60)
        self.cmb_grid.setCurrentText("1.0")
        self.cmb_wait.setCurrentText("开始前等待")
        self.cmb_launch_mode.setCurrentText("按数量投车")
        self.cmb_seq.setCurrentText("顺排(A→B→C)")
        self.spn_max_run.setValue(10)

    def _collect_inputs(self):
        project = self.ed_project.text().strip() or "工程"
        cars_a = int(self.spn_a_cars.value())
        cars_b = int(self.spn_b_cars.value())
        cars_c = int(self.spn_c_cars.value())
        analysis_minutes_for_ratio = int(self.spn_total_cars.value()) if hasattr(self, "spn_total_cars") else 0
        target_takt = float(self.spn_target_takt.value()) if hasattr(self, "spn_target_takt") else 0.0
        self.current_analysis_time_seconds = None
        self.current_theoretical_launch_count = None

        try:
            grid_step = float(self.cmb_grid.currentText())
            if grid_step <= 0:
                grid_step = 1.0
        except Exception:
            grid_step = 1.0
        wait_policy = "before"
        is_ratio_mode = self.cmb_launch_mode.currentIndex() == 1
        if is_ratio_mode:
            sequence_mode = "ratio"
        elif self.cmb_seq.currentIndex() == 1:
            sequence_mode = "alternate"
        else:
            sequence_mode = "grouped"
        max_consecutive = int(self.spn_max_run.value())

        ratio_pattern = None

        if sequence_mode == "ratio":
            ratio_vals = [cars_a, cars_b, cars_c]
            if any(v < 0 for v in ratio_vals) or sum(ratio_vals) <= 0:
                raise ValueError("按比例投车模式下，请填写 A/B/C 比例，且比例合计必须大于 0。")
            if analysis_minutes_for_ratio <= 0:
                raise ValueError("按比例投车模式下，请填写分析时间，且分析时间必须大于 0 分钟。")
            if target_takt <= 0:
                raise ValueError("按比例投车模式下，目标节拍必须大于 0，才能按分析时间计算理论投车台数。")

            analysis_time_seconds = analysis_minutes_for_ratio * 60.0
            theoretical_launch_count = int(math.ceil(analysis_time_seconds / target_takt)) + 50
            if theoretical_launch_count <= 0:
                raise ValueError("分析时间过短，按当前目标节拍计算的投车生成台数为 0，请增加分析时间。")

            ratio_pattern = {
                "A": ratio_vals[0],
                "B": ratio_vals[1],
                "C": ratio_vals[2],
            }
            cars = theoretical_launch_count
            self.current_analysis_time_seconds = analysis_time_seconds
            self.current_theoretical_launch_count = theoretical_launch_count
        else:
            cars = cars_a + cars_b + cars_c
            if cars <= 0:
                raise ValueError("A/B/C 车型数量合计必须大于 0")

        defs = []
        for r in range(self.tbl.rowCount()):
            seq = (self.tbl.item(r, 0).text().strip() if self.tbl.item(r, 0) else "")
            name = (self.tbl.item(r, 1).text().strip() if self.tbl.item(r, 1) else "")

            device_count_widget = self.tbl.cellWidget(r, 2)
            device_count_text = device_count_widget.currentText().strip() if isinstance(device_count_widget, QComboBox) else "2"

            line_scope_widget = self.tbl.cellWidget(r, 3)
            line_scope = line_scope_widget.currentText().strip() if isinstance(line_scope_widget, QComboBox) else "双线"

            grp = (self.tbl.item(r, 4).text().strip() if self.tbl.item(r, 4) else "")
            dur_a = (self.tbl.item(r, 5).text().strip() if self.tbl.item(r, 5) else "")
            dur_b = (self.tbl.item(r, 6).text().strip() if self.tbl.item(r, 6) else "")
            dur_c = (self.tbl.item(r, 7).text().strip() if self.tbl.item(r, 7) else "")
            color_hex = ""

            if not name and not grp and not dur_a and not dur_b and not dur_c:
                continue
            if not name or not grp:
                raise ValueError(f"第 {r + 1} 行请填写工程名称和岗位设备。")

            def _to_float(txt: str, field_name: str):
                try:
                    return float(txt)
                except Exception:
                    raise ValueError(f"第 {r + 1} 行『{field_name}』不是有效数字：{txt}")

            active_by_type = {
                "A": cars_a > 0 if not is_ratio_mode else cars_a > 0,
                "B": cars_b > 0 if not is_ratio_mode else cars_b > 0,
                "C": cars_c > 0 if not is_ratio_mode else cars_c > 0,
            }

            def _parse_duration(txt: str, car_type: str):
                if txt == "":
                    if active_by_type.get(car_type, False):
                        raise ValueError(f"第 {r + 1} 行『{car_type}工时』不能为空；参与投车车型若跳过该岗位请填写 0。")
                    return 0.0
                return _to_float(txt, f"{car_type}工时")

            duration_a = _parse_duration(dur_a, "A")
            duration_b = _parse_duration(dur_b, "B")
            duration_c = _parse_duration(dur_c, "C")

            try:
                device_count = int(device_count_text)
            except Exception:
                raise ValueError(f"第 {r + 1} 行『设备数量』不是有效值：{device_count_text}")

            if device_count not in (1, 2):
                raise ValueError(f"第 {r + 1} 行『设备数量』当前仅支持 1 或 2。")

            if device_count == 2 and line_scope != "双线":
                raise ValueError(f"第 {r + 1} 行设备数量为 2 时，所属线别必须为『双线』。")
            if device_count == 1 and line_scope == "双线":
                raise ValueError(f"第 {r + 1} 行设备数量为 1 时，所属线别不能为『双线』，请选择 1号线 / 2号线 / 双线共用。")

            if active_by_type.get("A", False) and active_by_type.get("B", False) and active_by_type.get("C", False):
                pass

            if (
                (not active_by_type.get("A", False) or duration_a == 0)
                and (not active_by_type.get("B", False) or duration_b == 0)
                and (not active_by_type.get("C", False) or duration_c == 0)
            ):
                raise ValueError(f"第 {r + 1} 行参与投车的车型不能全部跳过该岗位，请至少填写一个大于 0 的工时。")

            capacity = 2 if device_count == 2 else 1
            if device_count == 2:
                run_mode = "双线双设备"
            elif line_scope == "双线共用":
                run_mode = "双线单设备"
            else:
                run_mode = "单线单设备"

            rec = {
                "seq": int(float(seq)) if seq else len(defs) + 1,
                "display": name,
                "group": grp,
                "capacity": capacity,
                "durations": [duration_a],
                "color": color_hex,
                "run_mode": run_mode,
                "device_count": device_count,
                "line_scope": line_scope,
                "duration_a": duration_a,
                "duration_b": duration_b,
                "duration_c": duration_c,
            }
            defs.append(rec)

        defs.sort(key=lambda x: x["seq"])
        if not defs:
            raise ValueError("请至少填写一行有效的步骤（工程名称/岗位设备/A工时）")

        vehicle_counts = {
            "A": cars_a,
            "B": cars_b,
            "C": cars_c,
        }

        return project, cars, grid_step, wait_policy, defs, vehicle_counts, sequence_mode, max_consecutive, ratio_pattern, target_takt

    def do_export(self):
        try:
            project, cars, grid_step, wait_policy, defs, vehicle_counts, sequence_mode, max_consecutive, ratio_pattern, target_takt = self._collect_inputs()
        except Exception as e:
            QMessageBox.warning(self, "输入有误", str(e))
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出位置",
            f"{project}_组合票.xlsx",
            "Excel (*.xlsx)",
        )
        if not path:
            return
        self.dst_path = path

        worker = Worker(
            tickets.schedule_and_export,
            defs, cars, grid_step, wait_policy, project, self.dst_path,
            vehicle_counts, sequence_mode, max_consecutive, ratio_pattern, target_takt,
        )
        worker.signals.error.connect(self._on_error)
        worker.signals.finished.connect(self._on_export_finished)
        self.thread_pool.start(worker)
        self.status.showMessage("正在生成组合票...", 5000)

    def _on_export_finished(self, *args):
        self.status.showMessage("导出完成", 6000)
        QMessageBox.information(self, "完成", f"已导出：\n{self.dst_path}")

    def _fmt_analysis_num(self, value):
        try:
            v = float(value)
        except Exception:
            return str(value)
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        return f"{v:.1f}"
    
    def _apply_time_window_analysis(self, analysis, rows, target_takt):
        """v2-6：按比例投车的时间窗口产能分析。

        最终收口口径：
        - 计划下线台数 = 空线起步口径下，规定时间内按目标节拍理论可下线的整车台数。
        - 真实实际下线台数 = 分析时间内实际完成全部工程的车辆数量。
        - 显示实际下线台数 = min(真实实际下线台数, 计划下线台数)。
        - 达成率 = min(真实实际下线台数 / 计划下线台数, 100%)。
        - 实际下线节拍 = 用计划第 N 台理论下线时间与实际第 N 台下线时间差异反推。
        - 最终判定：真实实际下线台数达到计划下线台数，且实际第 N 台不晚于计划第 N 台时为 OK，否则 NG。
        - 累计阻塞、超节拍工程只显示，不直接参与最终判定。
        """
        if not isinstance(analysis, dict):
            return analysis

        analysis_time_seconds = getattr(self, "current_analysis_time_seconds", None)
        theoretical_launch_count = getattr(self, "current_theoretical_launch_count", None)
        if not analysis_time_seconds or not theoretical_launch_count or target_takt <= 0:
            return analysis

        summary = analysis.setdefault("summary", {})

        car_finish_times = {}
        station_names = []
        seen_stations = set()
        for row in rows or []:
            station = str(row.get("step_display", row.get("station", row.get("group", ""))) or "")
            if station and station not in seen_stations:
                station_names.append(station)
                seen_stations.add(station)

            try:
                car = int(row.get("car", 0) or 0)
            except Exception:
                car = 0

            if car <= 0:
                continue

            finish = 0.0
            for key in ("depart", "end", "svc_finish"):
                try:
                    finish = max(finish, float(row.get(key, 0.0) or 0.0))
                except Exception:
                    pass

            car_finish_times[car] = max(car_finish_times.get(car, 0.0), finish)

        finish_times = sorted(car_finish_times.values())

        actual_output_count = sum(
            1 for finish in finish_times
            if finish <= float(analysis_time_seconds) + 1e-9
        )

        station_count = max(1, len(station_names))
        line_lead_time = station_count * float(target_takt)
        if float(analysis_time_seconds) + 1e-9 < line_lead_time:
            planned_output_count = 0
        else:
            planned_output_count = math.floor(
                (float(analysis_time_seconds) - line_lead_time) / float(target_takt)
            ) + 1

        if planned_output_count > 0:
            display_actual_output_count = min(actual_output_count, planned_output_count)
        else:
            display_actual_output_count = actual_output_count

        planned_n_finish_time = None
        actual_n_finish_time = None
        finish_delta = None
        actual_line_takt = None

        if planned_output_count <= 0:
            achievement_rate = 0.0
            final_result = "未判定"
        else:
            achievement_rate = min(actual_output_count / planned_output_count, 1.0)
            planned_n_finish_time = line_lead_time + (planned_output_count - 1) * float(target_takt)

            if len(finish_times) >= planned_output_count:
                actual_n_finish_time = finish_times[planned_output_count - 1]
                finish_delta = actual_n_finish_time - planned_n_finish_time
                if planned_output_count > 1:
                    actual_line_takt = float(target_takt) + finish_delta / (planned_output_count - 1)
                else:
                    actual_line_takt = float(target_takt)

            if actual_output_count < planned_output_count:
                final_result = "NG"
            elif actual_n_finish_time is None:
                final_result = "NG"
            elif actual_n_finish_time > planned_n_finish_time + 1e-9:
                final_result = "NG"
            else:
                final_result = "OK"

        summary.update({
            "analysis_time_seconds": analysis_time_seconds,
            "analysis_time_minutes": analysis_time_seconds / 60.0,
            "theoretical_launch_count": int(theoretical_launch_count),
            "station_count": station_count,
            "line_lead_time": line_lead_time,
            "planned_output_count_in_window": planned_output_count,
            "actual_output_count_in_window": actual_output_count,
            "display_actual_output_count_in_window": display_actual_output_count,
            "actual_equivalent_count_in_window": display_actual_output_count,
            "achievement_rate": achievement_rate,
            "planned_n_finish_time": planned_n_finish_time,
            "actual_n_finish_time": actual_n_finish_time,
            "finish_delta": finish_delta,
            "actual_line_takt_in_window": actual_line_takt,
            "actual_production_takt_in_window": actual_line_takt,
            "time_window_result": final_result,
        })

        summary.pop("ok_output_count_in_window", None)
        summary.pop("output_gap_count", None)
        summary.pop("output_gap_time", None)
        summary.pop("actual_output_takt_in_window", None)
        summary.pop("completed_step_count_in_window", None)
        summary.pop("planned_step_count_in_window", None)
        summary.pop("planned_equivalent_count_raw_in_window", None)
        summary.pop("actual_equivalent_count_raw_in_window", None)
        summary.pop("entered_step_count_in_window", None)
        summary.pop("actual_takt_in_window", None)

        return analysis

    def _show_analysis_result(self, analysis):
        summary = analysis.get("summary", {}) if isinstance(analysis, dict) else {}

        total_cars = summary.get("total_cars", 0)
        max_finish = self._fmt_analysis_num(summary.get("max_finish", 0.0))
        total_wait = self._fmt_analysis_num(summary.get("total_wait", 0.0))
        avg_wait = self._fmt_analysis_num(summary.get("avg_wait", 0.0))
        blocking_result = summary.get("blocking_result", "无阻塞") or "无阻塞"
        blocking_time = self._fmt_analysis_num(summary.get("total_wait", 0.0))
        overflow_vehicle_count = self._fmt_analysis_num(summary.get("overflow_vehicle_count", 0.0))
        blocking_station_text = summary.get("blocking_station_text", "无") or "无"
        batch_overrun_time = self._fmt_analysis_num(summary.get("batch_overrun_time", 0.0))
        batch_overrun_cars = self._fmt_analysis_num(summary.get("batch_overrun_cars", 0.0))
        process_root_text = summary.get("process_over_takt_root_text", "无") or "无"
        takt_result = summary.get("takt_result", "未设定") or "未设定"
        over_count = summary.get("over_takt_station_count", 0)
        process_root_text = summary.get("process_over_takt_root_text", "无") or "无"
        batch_overrun_raw = float(summary.get("batch_overrun_time", 0.0) or 0.0)
        final_result = "NG" if batch_overrun_raw > 0 else "OK"

        text = (
            "结果分析："
            f"模型最终判定 {final_result} ｜ "
            f"累计阻塞 {blocking_time} 秒 ｜ "
            f"溢出工时 {batch_overrun_time} 秒 / {batch_overrun_cars} 台 ｜ "
            f"超节拍工程 {process_root_text}"
        )


        if hasattr(self, "lbl_analysis"):
            self.lbl_analysis.setText(text)
        else:
            QMessageBox.information(self, "排程分析完成", text)

        if hasattr(self, "_show_station_analysis"):
            self._show_station_analysis(analysis)

        if hasattr(self, "_show_vehicle_summary"):
            self._show_vehicle_summary(analysis)


    def _on_error(self, err_msg):
        self.status.showMessage("导出失败", 6000)
        QMessageBox.critical(self, "导出失败", str(err_msg))


    # ---------- 帮助弹窗 ----------
    def show_help(self):
        msg = (
            "<h3>组合票操作指南</h3>"
            "<ol>"
            "<li>多工程组合票按『设备数量 + 所属线别 + 岗位设备 + A/B/C 工时』录入</li>"
            "<li>设备数量可选：1 / 2；设备数量为 2 时，所属线别固定为『双线』</li>"
            "<li>设备数量为 1 时，所属线别可选：1号线 / 2号线 / 双线共用</li>"
            "<li>A/B/C 工时大于 0 表示该车型经过该岗位；工时为 0 表示该车型跳过该岗位</li>"
            "<li>参与投车的车型，工时不能为空；未参与投车的车型，工时可以为空</li>"
            "<li>投车模式支持：按数量投车 / 按比例投车；按数量投车下可选择顺排(A→B→C)或交替混流</li>"
            "<li>顺排/交替混流模式下，A/B/C 填数量；按比例投车模式下，A/B/C 填比例，并用分析时间与目标节拍计算理论投车台数</li>"
            "<li>最大连续台数默认 10；填 1 表示尽量强制交替</li>"
            "<li>填写完点击『分析当前排程』可查看结果；点击『生成并导出组合票』即可生成 Excel</li>"
            "</ol>"
        )
        QMessageBox.information(self, "帮助", msg)


    def _show_station_analysis(self, analysis):
        if not hasattr(self, "tbl_station_analysis"):
            return

        station_stats = analysis.get("station_stats", []) if isinstance(analysis, dict) else []
        self.tbl_station_analysis.setRowCount(0)

        for item in station_stats:
            r = self.tbl_station_analysis.rowCount()
            self.tbl_station_analysis.insertRow(r)

            values = [
                str(item.get("station", "")),
                str(item.get("count", 0)),
                self._fmt_analysis_num(item.get("total_process", 0.0)),
                self._fmt_analysis_num(item.get("blocking_time", item.get("overflow_wait_time", item.get("total_block_wait", 0.0)))),
                self._fmt_analysis_num(item.get("avg_process", 0.0)),
                self._fmt_analysis_num(item.get("avg_overflow_wait", item.get("avg_block_wait", 0.0))),
                str(item.get("takt_result", "未设定")),
                str(item.get("over_takt_types", "—")),
            ]

            for c, value in enumerate(values):
                table_item = QTableWidgetItem(value)
                if c > 0:
                    table_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    table_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.tbl_station_analysis.setItem(r, c, table_item)

        # 固定列宽，不再自动根据内容调整，避免每次分析后表格宽度跳动
        # self.tbl_station_analysis.resizeColumnsToContents()
    def _show_vehicle_summary(self, analysis):
        """在车型数据摘要区展示当前排程的车型构成与关键结果。"""
        if not hasattr(self, "lbl_vehicle_summary"):
            return

        summary = analysis.get("summary", {}) if isinstance(analysis, dict) else {}
        car_stats = analysis.get("car_stats", []) if isinstance(analysis, dict) else []
        car_type_summary = analysis.get("car_type_summary", analysis.get("type_stats", [])) if isinstance(analysis, dict) else []

        type_counts = {"A": 0, "B": 0, "C": 0}
        if car_stats:
            for item in car_stats:
                vt = str(item.get("car_type", "") or "").strip().upper()
                if vt in type_counts:
                    type_counts[vt] += 1
        else:
            for item in car_type_summary:
                vt = str(item.get("car_type", "") or "").strip().upper()
                if vt in type_counts:
                    type_counts[vt] += int(item.get("count", 0) or 0)

        total_cars = int(summary.get("total_cars", 0) or 0)
        if total_cars <= 0:
            total_cars = sum(type_counts.values())

        # v2-5 blocking analysis fields
        target_takt = self._fmt_analysis_num(summary.get("target_takt", 0.0))
        max_finish = self._fmt_analysis_num(summary.get("max_finish", 0.0))
        total_wait = self._fmt_analysis_num(summary.get("total_wait", 0.0))
        avg_wait = self._fmt_analysis_num(summary.get("avg_wait", 0.0))
        takt_result = summary.get("takt_result", "未设定") or "未设定"
        over_count = summary.get("over_takt_station_count", 0)
        blocking_result = summary.get("blocking_result", "无阻塞") or "无阻塞"
        blocking_station_text = summary.get("blocking_station_text", "无") or "无"
        blocking_station_count = summary.get("blocking_station_count", 0)
        blocking_time = self._fmt_analysis_num(summary.get("total_wait", 0.0))
        overflow_vehicle_count = self._fmt_analysis_num(summary.get("overflow_vehicle_count", 0.0))
        batch_overrun_time = self._fmt_analysis_num(summary.get("batch_overrun_time", 0.0))
        batch_overrun_cars = self._fmt_analysis_num(summary.get("batch_overrun_cars", 0.0))
        process_root_text = summary.get("process_over_takt_root_text", "无") or "无"

        batch_overrun_raw = float(summary.get("batch_overrun_time", 0.0) or 0.0)

        time_window_result = summary.get("time_window_result", "") or ""
        final_result = time_window_result or "OK"

        actual_output_count = int(summary.get("actual_output_count_in_window", 0) or 0)
        planned_output_count = self._fmt_analysis_num(summary.get("planned_output_count_in_window", 0.0))
        theoretical_launch_count = int(summary.get("theoretical_launch_count", total_cars) or total_cars)
        achievement_rate = float(summary.get("achievement_rate", 0.0) or 0.0) * 100
        target_takt_display = self._fmt_analysis_num(summary.get("target_takt", self.spn_target_takt.value() if hasattr(self, "spn_target_takt") else 0))

        def _fmt_optional_seconds(value):
            if value is None:
                return "—"
            return self._fmt_analysis_num(value)

        planned_n_finish_time = _fmt_optional_seconds(summary.get("planned_n_finish_time"))
        actual_n_finish_time = _fmt_optional_seconds(summary.get("actual_n_finish_time"))
        finish_delta_raw = summary.get("finish_delta")
        if finish_delta_raw is None:
            finish_delta = "—"
        else:
            try:
                finish_delta_num = float(finish_delta_raw)
                finish_delta = f"{finish_delta_num:+.1f}" if abs(finish_delta_num - round(finish_delta_num)) >= 1e-9 else f"{int(round(finish_delta_num)):+d}"
            except Exception:
                finish_delta = str(finish_delta_raw)
        actual_line_takt_in_window = _fmt_optional_seconds(summary.get("actual_line_takt_in_window"))

        is_ratio = hasattr(self, "cmb_launch_mode") and self.cmb_launch_mode.currentIndex() == 1
        if is_ratio:
            ratio_a = int(self.spn_a_cars.value()) if hasattr(self, "spn_a_cars") else 0
            ratio_b = int(self.spn_b_cars.value()) if hasattr(self, "spn_b_cars") else 0
            ratio_c = int(self.spn_c_cars.value()) if hasattr(self, "spn_c_cars") else 0
            analysis_minutes = int(self.spn_total_cars.value()) if hasattr(self, "spn_total_cars") else 0
            mode_line = (
                f"投车模式：按比例投车，A:B:C = {ratio_a}:{ratio_b}:{ratio_c}，"
                f"分析时间 {analysis_minutes} 分钟。"
            )
        else:
            seq_text = self.cmb_seq.currentText() if hasattr(self, "cmb_seq") else "—"
            mode_line = f"投车模式：按数量投车，排列方式：{seq_text}。"

        if is_ratio:
            left_html = (
                "<b>基础排程</b><br>"
                "计时参考：首车进入首岗位时刻为 0s<br>"
                f"{mode_line}<br>"
                f"车型生成：A{type_counts['A']}台 / B{type_counts['B']}台 / C{type_counts['C']}台（共{total_cars}台）"
            )
        else:
            left_html = (
                "<b>基础排程</b><br>"
                "计时参考：首车进入首岗位时刻为 0s<br>"
                f"{mode_line}<br>"
                f"实际生成：A {type_counts['A']} 台 / B {type_counts['B']} 台 / C {type_counts['C']} 台（共 {total_cars} 台）"
            )
        if is_ratio and time_window_result:
            right_html = (
                "<b>模型判定</b><br>"
                f"最终判定：{final_result}<br>"
                f"下线达成：计划{planned_output_count}台｜实际{actual_output_count}台｜达成率{achievement_rate:.1f}%<br>"
                f"计划完成：计划{planned_n_finish_time}s｜实际{actual_n_finish_time}s｜差异{finish_delta}s<br>"
                f"实际下线节拍：{actual_line_takt_in_window}s/台｜目标{target_takt_display}s/台<br>"
                f"阻塞分析：累计阻塞{blocking_time}s｜超节拍工程：{process_root_text}"
            )
        else:
            try:
                target_takt_raw = float(summary.get(
                    "target_takt",
                    self.spn_target_takt.value() if hasattr(self, "spn_target_takt") else 0.0
                ) or 0.0)
            except Exception:
                target_takt_raw = 0.0

            try:
                max_finish_raw = float(summary.get("max_finish", 0.0) or 0.0)
            except Exception:
                max_finish_raw = 0.0

            try:
                total_wait_raw = float(summary.get("total_wait", 0.0) or 0.0)
            except Exception:
                total_wait_raw = 0.0

            station_count = 0
            try:
                station_count = len(self._sim_station_names())
            except Exception:
                station_count = 0

            if station_count <= 0:
                station_stats = analysis.get("station_stats", []) if isinstance(analysis, dict) else []
                station_count = len(station_stats)

            station_count = max(1, station_count)

            if target_takt_raw > 0 and total_cars > 0:
                planned_finish_raw = (total_cars + station_count - 1) * target_takt_raw
            else:
                planned_finish_raw = 0.0

            planned_finish = self._fmt_analysis_num(planned_finish_raw)

            has_over_takt_root = process_root_text not in ("无", "", "—")
            has_finish_gap = planned_finish_raw > 0 and max_finish_raw > planned_finish_raw + 1e-6
            has_blocking = total_wait_raw > 0

            quantity_final_result = "NG" if (has_finish_gap or has_blocking or has_over_takt_root) else "OK"

            right_html = (
                "<b>模型判定</b><br>"
                f"最终判定：{quantity_final_result}<br>"
                f"计划完成时间：{planned_finish}s｜总完成时间：{max_finish}s<br>"
                f"累计阻塞：{blocking_time}s｜超节拍工程：{process_root_text}"
            )
        
        html = (
            "<table width='100%' cellspacing='0' cellpadding='0'>"
            "<tr>"
            f"<td width='50%' valign='top' style='padding-right:12px;'>{left_html}</td>"
            f"<td width='50%' valign='top' style='padding-left:12px; border-left:1px solid #cbd5e1;'>{right_html}</td>"
            "</tr>"
            "</table>"
        )
        self.lbl_vehicle_summary.setText(html)
    def _active_sim_rows(self):
        """返回当前仿真时间正在加工的排程段。"""
        rows = getattr(self, "last_schedule_rows", []) or []
        current = float(getattr(self, "sim_time", 0.0) or 0.0)
        active = []
        for row in rows:
            start = self._sim_row_start(row)
            end = self._sim_row_end(row)
            if start <= current < end:
                active.append((start, end, row))
        active.sort(key=lambda x: (self._sim_row_station(x[2]), x[0]))
        return active

    def _draw_sim_scene(self):
        """绘制生产线轨道沙盘风仿真：固定岗位卡片 + 轨道线 + PROC/WAIT 车辆。"""
        if not hasattr(self, "sim_scene"):
            return

        self.sim_scene.clear()
        rows = getattr(self, "last_schedule_rows", []) or []
        if not rows:
            self.sim_scene.addText("请先点击『分析当前排程』生成仿真数据。")
            return

        station_names = self._sim_station_names()
        if not station_names:
            self.sim_scene.addText("暂无岗位数据。")
            return

        current = float(getattr(self, "sim_time", 0.0) or 0.0)
        total = float(getattr(self, "last_max_finish", 0.0) or 0.0)

        compact_mode = len(station_names) > 6
        margin_x = 20 if compact_mode else 24
        station_w = 96 if compact_mode else 132
        station_h = 54 if compact_mode else 60
        gap = 12 if compact_mode else 24
        y_station = 34
        y_track = y_station + station_h + 34
        y_car_base = y_track + 16
        car_w = 54 if compact_mode else 72
        car_h = 28 if compact_mode else 34

        scene_w = (
            margin_x * 2
            + len(station_names) * station_w
            + max(0, len(station_names) - 1) * gap
            + 120
        )
        scene_h = y_car_base + 4 * (car_h + 10) + 36
        scene_w = max(scene_w, 760)
        scene_h = max(scene_h, 270 if compact_mode else 250)

        # 浅色工业沙盘背景
        self.sim_scene.addRect(
            0,
            0,
            scene_w,
            scene_h,
            QPen(QColor("#f1f5f9")),
            QBrush(QColor("#f1f5f9")),
        )

        title = self.sim_scene.addText(f"SIM TIME：{current:.1f}s / {total:.1f}s")
        title.setDefaultTextColor(QColor("#334155"))
        title.setPos(margin_x, 6)

        # 极简图例
        legend = self.sim_scene.addText("A 蓝｜B 橙｜C 绿｜PROC 加工｜WAIT 等待｜IDLE 空闲")
        legend.setDefaultTextColor(QColor("#64748b"))
        legend.setPos(margin_x + 220, 6)

        station_x = {}
        station_centers = []
        for idx, name in enumerate(station_names):
            x = margin_x + idx * (station_w + gap)
            station_x[name] = x
            station_centers.append(x + station_w / 2)

            self.sim_scene.addRect(
                x,
                y_station,
                station_w,
                station_h,
                QPen(QColor("#cbd5e1"), 1),
                QBrush(QColor("#ffffff")),
            )

            station_label = f"ST-{idx + 1}\n{name}"
            text = self.sim_scene.addText(station_label)
            text.setDefaultTextColor(QColor("#1e293b"))
            text.setPos(x + 7, y_station + 5)

        # 主轨道线：岗位固定，车辆在轨道层活动
        if station_centers:
            track_start = station_centers[0]
            track_end = station_centers[-1]
            self.sim_scene.addLine(
                track_start,
                y_track,
                track_end,
                y_track,
                QPen(QColor("#94a3b8"), 3),
            )
            for center in station_centers:
                self.sim_scene.addEllipse(
                    center - 4,
                    y_track - 4,
                    8,
                    8,
                    QPen(QColor("#64748b")),
                    QBrush(QColor("#ffffff")),
                )

        car_draw_items = []
        target_takt = float(self.spn_target_takt.value()) if hasattr(self, "spn_target_takt") else 0.0
        station_proc_counts = {name: 0 for name in station_names}
        station_wait_counts = {name: 0 for name in station_names}
        station_over_counts = {name: 0 for name in station_names}

        car_rows = self._sim_car_rows()
        for _, car_segments in car_rows.items():
            if not car_segments:
                continue

            first_start = self._sim_row_start(car_segments[0])
            last_end = self._sim_row_end(car_segments[-1])

            # 未进入排程、已完成车辆都不显示，保持画面聚焦当前状态
            if current < first_start or current >= last_end:
                continue

            label = self._sim_row_car_label(car_segments[0])
            car_type = str(
                self._sim_row_value(
                    car_segments[0],
                    "car_type",
                    "type",
                    "vehicle_type",
                    default=""
                ) or ""
            ).upper()

            for i, seg in enumerate(car_segments):
                start = self._sim_row_start(seg)
                end = self._sim_row_end(seg)
                station = self._sim_row_station(seg)
                station_base_x = station_x.get(station, margin_x)
                next_seg = car_segments[i + 1] if i + 1 < len(car_segments) else None

                # PROC：显示在当前岗位下方，只在岗位范围内轻微移动，不提前跨到下一岗位
                if start <= current < end:
                    station_proc_counts[station] = station_proc_counts.get(station, 0) + 1
                    elapsed_in_segment = max(0.0, current - start)
                    is_over_takt = target_takt > 0 and elapsed_in_segment > target_takt
                    if is_over_takt:
                        station_over_counts[station] = station_over_counts.get(station, 0) + 1
                    progress = 0.0
                    if end > start:
                        progress = max(0.0, min(1.0, (current - start) / (end - start)))
                    inner_move = max(0, station_w - car_w - 4)
                    x_pos = station_base_x + 2 + progress * inner_move
                    car_draw_items.append({
                        "label": label,
                        "car_type": car_type,
                        "status": "PROC",
                        "station": station,
                        "x": x_pos,
                        "order_time": start,
                        "over_takt": is_over_takt,
                    })
                    break

                # WAIT：停在原岗位出口附近，表示等待进入下一岗位
                if next_seg is not None:
                    next_start = self._sim_row_start(next_seg)
                    if end <= current < next_start:
                        station_wait_counts[station] = station_wait_counts.get(station, 0) + 1
                        x_pos = station_base_x + station_w - car_w - 6
                        car_draw_items.append({
                            "label": label,
                            "car_type": car_type,
                            "status": "WAIT",
                            "station": station,
                            "x": x_pos,
                            "order_time": end,
                        })
                        break

        # Andon 状态提示：固定在岗位卡片底部，岗位与车辆分离但关系清晰
        for station in station_names:
            x = station_x.get(station, margin_x)
            proc_count = station_proc_counts.get(station, 0)
            wait_count = station_wait_counts.get(station, 0)
            over_count = station_over_counts.get(station, 0)
            # 岗位状态只用颜色提示：红=超节拍，橙=等待，绿=加工，灰=空闲
            status_text = "●"
            if over_count > 0:
                status_color = QColor("#dc2626")
                card_border = QColor("#dc2626")
            elif wait_count > 0:
                status_color = QColor("#f97316")
                card_border = QColor("#f97316")
            elif proc_count > 0:
                status_color = QColor("#16a34a")
                card_border = QColor("#16a34a")
            else:
                status_color = QColor("#94a3b8")
                card_border = QColor("#cbd5e1")

            # 状态色边条，强化工位当前状态
            self.sim_scene.addRect(
                x,
                y_station,
                4,
                station_h,
                QPen(card_border),
                QBrush(card_border),
            )

            status_item = self.sim_scene.addText(status_text)
            status_item.setDefaultTextColor(status_color)
            status_item.setPos(x + 7, y_station + station_h - 19)

        # 车辆块：只显示 PROC / WAIT，不显示 DONE
        status_order = {"PROC": 0, "WAIT": 1}
        car_draw_items.sort(
            key=lambda item: (
                station_names.index(item["station"]) if item["station"] in station_names else 999,
                status_order.get(item["status"], 9),
                item["order_time"],
            )
        )

        lane_counts = {}
        max_lane = 0
        for item in car_draw_items[:24]:
            station = item["station"]
            lane = lane_counts.get(station, 0)
            lane_counts[station] = lane + 1
            max_lane = max(max_lane, lane + 1)

            x = float(item["x"])
            y = y_car_base + lane * (car_h + 10)
            status = item["status"]
            car_type = item["car_type"]

            if item.get("over_takt"):
                border = QColor("#dc2626")
            elif status == "PROC":
                border = QColor("#16a34a")
            elif status == "WAIT":
                border = QColor("#f97316")
            else:
                border = QColor("#94a3b8")

            if car_type == "A":
                fill = QColor("#dbeafe")
                type_bar = QColor("#2563eb")
            elif car_type == "B":
                fill = QColor("#ffedd5")
                type_bar = QColor("#ea580c")
            elif car_type == "C":
                fill = QColor("#dcfce7")
                type_bar = QColor("#16a34a")
            else:
                fill = QColor("#f1f5f9")
                type_bar = QColor("#64748b")

            # 工业仿真小车块：阴影 + 彩色车身 + 大编号 + 双轮，适合小尺寸动画快速识别
            body_h = car_h * 0.66
            body_y = y + car_h * 0.14
            shadow_offset = 3

            self.sim_scene.addRect(
                x + shadow_offset,
                body_y + shadow_offset,
                car_w,
                body_h,
                QPen(QColor("#cbd5e1"), 1),
                QBrush(QColor("#cbd5e1")),
            )
            self.sim_scene.addRect(
                x,
                body_y,
                car_w,
                body_h,
                QPen(border, 2),
                QBrush(fill),
            )

            # 车型色条：左侧小色块，A/B/C 一眼区分
            self.sim_scene.addRect(
                x,
                body_y,
                5,
                body_h,
                QPen(type_bar),
                QBrush(type_bar),
            )

            # 前窗/车头提示：右侧浅色窗块，让方块仍有车辆方向感
            nose_w = 8 if compact_mode else 10
            self.sim_scene.addRect(
                x + car_w - nose_w - 3,
                body_y + 4,
                nose_w,
                max(8, body_h - 8),
                QPen(border, 1),
                QBrush(QColor("#f8fafc")),
            )

            # 顶部高光，增加一点立体感
            self.sim_scene.addLine(
                x + 8,
                body_y + 4,
                x + car_w - nose_w - 8,
                body_y + 4,
                QPen(QColor("#ffffff"), 1),
            )
            
            # 双轮：简化黑色轮胎
            wheel_r = 4 if compact_mode else 5
            wheel_y = body_y + body_h - 1
            front_wheel_x = x + car_w * 0.24
            rear_wheel_x = x + car_w * 0.76
            for wx in (front_wheel_x, rear_wheel_x):
                self.sim_scene.addEllipse(
                    wx - wheel_r,
                    wheel_y - wheel_r,
                    wheel_r * 2,
                    wheel_r * 2,
                    QPen(QColor("#1e293b"), 1),
                    QBrush(QColor("#1e293b")),
                )

            # 状态标签：超节拍/等待/加工时显示在车辆上方，车身内只放编号
            if item.get("over_takt"):
                tag_text = "超节拍"
                tag_color = QColor("#dc2626")
            elif status == "WAIT":
                tag_text = "WAIT"
                tag_color = QColor("#f97316")
            else:
                tag_text = "PROC"
                tag_color = QColor("#16a34a")

            tag_x = x
            tag_y = max(y - 2, body_y - 19)
            tag_bg_w = 44 if compact_mode else 52
            tag_bg_h = 16

            self.sim_scene.addRect(
                tag_x,
                tag_y,
                tag_bg_w,
                tag_bg_h,
                QPen(tag_color),
                QBrush(tag_color),
            )
            
            tag_item = self.sim_scene.addText(tag_text)
            tag_item.setDefaultTextColor(QColor("#ffffff"))
            tag_item.setPos(tag_x + 3, tag_y - 1)
            tag_item.setZValue(2)

            short_label = str(item["label"]).replace("Car#", "#").replace("(", "").replace(")", "")
            car_no = "".join(ch for ch in short_label if ch.isdigit()) or short_label

            car_text = self.sim_scene.addText(car_no)
            car_text.setDefaultTextColor(QColor("#0f172a"))
            car_text.setPos(x + car_w * 0.42, body_y + body_h * 0.23)
            car_text.setZValue(3)

        if not car_draw_items:
            empty = self.sim_scene.addText("当前时间无车辆处于加工或等待状态")
            empty.setDefaultTextColor(QColor("#64748b"))
            empty.setPos(margin_x, y_car_base)

        scene_h = y_car_base + max(1, max_lane) * (car_h + 10) + 42
        self.sim_scene.setSceneRect(0, 0, scene_w, max(scene_h, 270 if compact_mode else 250))
