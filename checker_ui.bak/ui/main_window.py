from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableView,
    QFileDialog, QPushButton, QLabel, QStatusBar, QMessageBox,
    QSplitter, QToolBar, QHeaderView, QSizePolicy
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QThreadPool, QTimer

from ..models.state import AppState
from ..models.dataframe_model import DataFrameModel
from ..infra.threads import Worker

from ..core import loaders as loaders, comparator as comparator, exporter as exporter


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("生产指示点检系统 (PySide6)")
        self.resize(1280, 800)

        self.state = AppState()
        self.thread_pool = QThreadPool.globalInstance()

        self._build_ui()
        self._connect_signals()

        # 仅做一次自动列宽的标记，避免频繁触发
        self._sized_std_once = False
        self._sized_sys_once = False

    # ---------------------- UI ---------------------- #
    def _build_ui(self):
        # 工具栏
        tb = QToolBar("Main")
        self.addToolBar(tb)

        # 返回主页
        self.act_back = QAction("返回主页", self)
        tb.addAction(self.act_back)
        tb.addSeparator()

        # 基本动作
        self.act_open_std = QAction("打开标准文件", self)
        self.act_open_sys = QAction("打开系统文件", self)
        self.act_compare = QAction("一致性校对", self)
        self.act_export = QAction("导出Excel", self)
        # 新增：自适应列宽（一次）
        self.act_fit_cols = QAction("自适应列宽（一次）", self)

        tb.addAction(self.act_open_std)
        tb.addAction(self.act_open_sys)
        tb.addSeparator()
        tb.addAction(self.act_compare)
        tb.addSeparator()
        tb.addAction(self.act_export)
        tb.addSeparator()
        tb.addAction(self.act_fit_cols)

        # 中央区域
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        # summary bar（小徽章行）
        summary_bar = QWidget(self)
        summary_layout = QHBoxLayout(summary_bar)
        summary_layout.setContentsMargins(0, 0, 0, 0)
        summary_layout.setSpacing(6)
        summary_bar.setMaximumHeight(36)  # 限高，避免占竖向空间
        summary_bar.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        self.lbl_std_ok = self._make_chip("标准 OK: 0", "#E8F5E9")
        self.lbl_std_ng = self._make_chip("标准 NG: 0", "#FFEBEE")
        self.lbl_sys_ok = self._make_chip("系统 OK: 0", "#E8F5E9")
        self.lbl_sys_ng = self._make_chip("系统 NG: 0", "#FFEBEE")

        summary_layout.addWidget(self.lbl_std_ok)
        summary_layout.addWidget(self.lbl_std_ng)
        summary_layout.addWidget(self.lbl_sys_ok)
        summary_layout.addWidget(self.lbl_sys_ng)
        summary_layout.addStretch()
        layout.addWidget(summary_bar)

        # 左右两表
        self.table_std = QTableView()
        self.table_sys = QTableView()

        self.model_std = DataFrameModel(status_col="比对结果")
        self.model_sys = DataFrameModel(status_col="比对结果")

        self.table_std.setModel(self.model_std)
        self.table_sys.setModel(self.model_sys)

        # ⚠️ 关键：不要用 ResizeToContents（会全表扫描，卡顿）
        for tv in (self.table_std, self.table_sys):
            tv.setAlternatingRowColors(True)
            tv.setSelectionBehavior(QTableView.SelectRows)
            tv.verticalHeader().setVisible(False)  # 隐藏行号
            header = tv.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.Interactive)  # 改为交互式，避免每次全表扫描
            header.setDefaultSectionSize(120)
            header.setMinimumSectionSize(60)
            header.setStretchLastSection(False)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.table_std)
        splitter.addWidget(self.table_sys)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setChildrenCollapsible(False)
        layout.addWidget(splitter)

        # 底部状态栏
        self.status = QStatusBar()
        self.setStatusBar(self.status)

    def _make_chip(self, text: str, bg: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f"padding:2px 8px; border-radius:4px; background:{bg};")
        return lbl

    def _update_summary_chips(self, ok_std: int, ng_std: int, ok_sys: int, ng_sys: int):
        self.lbl_std_ok.setText(f"标准 OK: {ok_std}")
        self.lbl_std_ng.setText(f"标准 NG: {ng_std}")
        self.lbl_sys_ok.setText(f"系统 OK: {ok_sys}")
        self.lbl_sys_ng.setText(f"系统 NG: {ng_sys}")

    # —— 快速列宽：只扫描前若干行/列，限制最大宽 —— #
    def _autosize_columns_fast(self, view: QTableView, max_rows: int = 200,
                               max_cols: int = 80, max_width: int = 420):
        model = view.model()
        if model is None:
            return
        try:
            view.setUpdatesEnabled(False)
            fm = view.fontMetrics()
            cols = min(model.columnCount(), max_cols)
            rows = min(model.rowCount(), max_rows)
            for col in range(cols):
                # 先用表头长度做一个基准
                header_text = model.headerData(col, Qt.Horizontal, Qt.DisplayRole) or ""
                width = fm.horizontalAdvance(str(header_text)) + 24
                for row in range(rows):
                    idx = model.index(row, col)
                    data = model.data(idx, Qt.DisplayRole)
                    if data is None:
                        continue
                    w = fm.horizontalAdvance(str(data)) + 24
                    if w > width:
                        width = w
                        if width >= max_width:
                            break
                view.setColumnWidth(col, min(width, max_width))
        finally:
            view.setUpdatesEnabled(True)

    # —— 比对返回校验 & 完成后恢复 —— #
    def _validate_compare_result(self, result):
        """校验 compare() 返回值，确保为 (std_df, sys_df) 且包含'比对结果'列"""
        try:
            std_df, sys_df = result
        except Exception as e:
            raise ValueError(f"compare() 应返回2个DataFrame，实际得到：{type(result)}") from e
        for name, df in (("标准", std_df), ("系统", sys_df)):
            if not hasattr(df, "columns"):
                raise TypeError(f"{name}结果不是 DataFrame：{type(df)}")
            if "比对结果" not in df.columns:
                raise KeyError(f"{name}结果缺少 '比对结果' 列")
        return std_df, sys_df

    def _after_compare(self):
        self.act_compare.setEnabled(True)
        try:
            self.unsetCursor()
        except Exception:
            pass

    # ---------------------- 信号连接 ---------------------- #
    def _connect_signals(self):
        # 工具栏
        self.act_open_std.triggered.connect(self.load_std)
        self.act_open_sys.triggered.connect(self.load_sys)
        self.act_compare.triggered.connect(self.do_compare)
        self.act_export.triggered.connect(self.export_excel)
        self.act_back.triggered.connect(self.go_home)
        self.act_fit_cols.triggered.connect(
            lambda: (self._autosize_columns_fast(self.table_std),
                     self._autosize_columns_fast(self.table_sys))
        )

    # 便于“入口页 -> 传入路径”复用
    def load_std_path(self, path: str):
        if not path:
            return
        worker = Worker(loaders.load_std_df, path)
        worker.signals.result.connect(self._on_std_loaded)
        worker.signals.error.connect(self._on_error)
        self.thread_pool.start(worker)
        self.status.showMessage("正在读取标准文件...", 3000)

    def load_sys_path(self, path: str):
        if not path:
            return
        worker = Worker(loaders.load_sys_df, path)
        worker.signals.result.connect(self._on_sys_loaded)
        worker.signals.error.connect(self._on_error)
        self.thread_pool.start(worker)
        self.status.showMessage("正在读取系统文件...", 3000)

    # ---------------------- 动作 ---------------------- #
    def load_std(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择标准文件", "", "Excel (*.xlsx *.xls)")
        if not path:
            return
        worker = Worker(loaders.load_std_df, path)
        worker.signals.result.connect(self._on_std_loaded)
        worker.signals.error.connect(self._on_error)
        self.thread_pool.start(worker)
        self.status.showMessage("正在读取标准文件...", 3000)

    def _on_std_loaded(self, df):
        self.state.std_df = df
        self.model_std.setDataFrame(df)
        # 只做一次轻量自适应（避免每次都扫全表）
        if not self._sized_std_once:
            QTimer.singleShot(0, lambda: self._autosize_columns_fast(self.table_std))
            self._sized_std_once = True
        self.status.showMessage("标准文件读取完成", 5000)

    def load_sys(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择系统文件", "", "Excel (*.xlsx *.xls)")
        if not path:
            return
        worker = Worker(loaders.load_sys_df, path)
        worker.signals.result.connect(self._on_sys_loaded)
        worker.signals.error.connect(self._on_error)
        self.thread_pool.start(worker)
        self.status.showMessage("正在读取系统文件...", 3000)

    def _on_sys_loaded(self, df):
        self.state.sys_df = df
        self.model_sys.setDataFrame(df)
        if not self._sized_sys_once:
            QTimer.singleShot(0, lambda: self._autosize_columns_fast(self.table_sys))
            self._sized_sys_once = True
        self.status.showMessage("系统文件读取完成", 5000)

    def do_compare(self):
        if self.state.std_df is None or self.state.sys_df is None:
            QMessageBox.warning(self, "提示", "请先加载标准文件和系统文件")
            return
        # 防重复点击 & 忙碌指示
        self.act_compare.setEnabled(False)
        try:
            self.setCursor(Qt.BusyCursor)
        except Exception:
            pass

        worker = Worker(comparator.compare, self.state.std_df, self.state.sys_df)
        worker.signals.result.connect(self._on_compared)
        worker.signals.error.connect(self._on_error)
        worker.signals.finished.connect(self._after_compare)
        self.thread_pool.start(worker)
        self.status.showMessage("正在比对...", 3000)

    def _on_compared(self, result):
        # 统一校验 compare() 返回值
        try:
            std_df, sys_df = self._validate_compare_result(result)
        except Exception as e:
            import traceback, pprint
            tb = traceback.format_exc()
            detail = f"{e}\n\n原始返回：{pprint.pformat(result)[:500]}"
            self._on_error(detail + "\n\n" + tb)
            return

        self.state.result_df = std_df
        self.state.sys_df = sys_df

        self.model_std.setDataFrame(std_df)
        self.model_sys.setDataFrame(sys_df)

        # 轻量自适应列宽（一次）
        QTimer.singleShot(0, lambda: self._autosize_columns_fast(self.table_std))
        QTimer.singleShot(0, lambda: self._autosize_columns_fast(self.table_sys))

        try:
            ok_std = (std_df["比对结果"] == "OK").sum()
            ng_std = (std_df["比对结果"] == "NG").sum()
            ok_sys = (sys_df["比对结果"] == "OK").sum()
            ng_sys = (sys_df["比对结果"] == "NG").sum()
        except Exception:
            ok_std = ng_std = ok_sys = ng_sys = 0

        self._update_summary_chips(ok_std, ng_std, ok_sys, ng_sys)
        self.status.showMessage("比对完成", 5000)

    def export_excel(self):
        if self.state.result_df is None or self.state.sys_df is None:
            QMessageBox.warning(self, "提示", "请完成一次比对后再导出")
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出结果", "对比结果.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        worker = Worker(exporter.export, self.state.result_df, self.state.sys_df, path)
        worker.signals.error.connect(self._on_error)
        worker.signals.finished.connect(lambda: QMessageBox.information(self, "完成", "导出成功"))
        self.thread_pool.start(worker)
        self.status.showMessage("正在导出...", 3000)

    def go_home(self):
        # 优先使用显式传入的首页引用
        home = getattr(self, "home_window", None)
        if home is not None and hasattr(home, "show"):
            try:
                home.show()
            except Exception:
                pass
        else:
            # 回退：若以 parent 方式管理，则尝试显示父窗口
            parent = self.parent()
            if parent is not None and hasattr(parent, "show"):
                try:
                    parent.show()
                except Exception:
                    pass
        self.close()


    def _on_error(self, tb: str):
        QMessageBox.critical(self, "出错了", tb)
        self.status.showMessage("发生错误", 5000)