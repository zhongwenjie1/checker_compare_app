# -*- coding: utf-8 -*-
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QToolBar, QStatusBar, QMessageBox, QTableWidget,
    QTableWidgetItem, QSpinBox, QComboBox, QLineEdit, QColorDialog,
    QGraphicsScene, QGraphicsView, QGraphicsRectItem, QGraphicsTextItem, QGraphicsLineItem, QSplitter, QDialog
)
from PySide6.QtCore import Qt, QThreadPool, QPointF
from PySide6.QtGui import QAction, QColor, QPainter, QPainterPath, QPen, QBrush

try:
    from checker_ui.infra.threads import Worker
    from checker_ui.core import tickets
except Exception:
    from ..infra.threads import Worker
    from ..core import tickets


class ExportTicketWindow(QMainWindow):
    """
    导出组合票（独立于数据校对）
    步骤表字段：
      序号 / 工序显示名 / 工位组名 / 并行能力 / 时长(秒，可逗号) /
      区域ID(可选) / 区域容量(可选) / 起始需等区域ID(可选)
    说明：
      - 同一“区域ID”的一串连续步骤视为一个“阻塞区域（Zone）”，容量=同时允许几台车处于该区域。
      - “起始需等区域ID”用于上游工位：本工位本身不占用区域名额，但开工/放行必须等待该区域出现空位。
    """
    COL_COLOR = 8

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("导出组合票")
        self.resize(1120, 720)

        self.thread_pool = QThreadPool.globalInstance()
        self.dst_path = None

        self._build_ui()
        self._connect_signals()

    # ---------------- UI ---------------- #
    def _build_ui(self):
        tb = QToolBar("Ticket")
        self.addToolBar(tb)

        self.act_back = QAction("返回主页", self)
        tb.addAction(self.act_back)
        tb.addSeparator()

        self.act_help = QAction("帮助", self)
        tb.addAction(self.act_help)
        tb.addSeparator()

        self.act_diagram = QAction("流程图", self)
        tb.addAction(self.act_diagram)
        tb.addSeparator()

        self.act_add_row = QAction("添加步骤", self)
        self.act_del_row = QAction("删除步骤", self)
        self.act_fill_sample = QAction("填入示例", self)
        tb.addAction(self.act_add_row)
        tb.addAction(self.act_del_row)
        tb.addSeparator()
        tb.addAction(self.act_fill_sample)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # 参数区
        row_top = QHBoxLayout()
        root.addLayout(row_top)

        row_top.addWidget(QLabel("工程名称："))
        self.ed_project = QLineEdit()
        self.ed_project.setPlaceholderText("例如：L2++")
        self.ed_project.setFixedWidth(220)
        row_top.addWidget(self.ed_project)

        row_top.addSpacing(12)
        row_top.addWidget(QLabel("车号数量："))
        self.spn_cars = QSpinBox()
        self.spn_cars.setRange(1, 9999)
        self.spn_cars.setValue(4)
        row_top.addWidget(self.spn_cars)

        row_top.addSpacing(12)
        row_top.addWidget(QLabel("时间格刻度："))
        self.cmb_grid = QComboBox()
        self.cmb_grid.addItems(["1.0", "0.5", "2.0"])
        self.cmb_grid.setCurrentIndex(0)
        row_top.addWidget(self.cmb_grid)

        row_top.addSpacing(12)
        row_top.addWidget(QLabel("等待分配："))
        self.cmb_wait = QComboBox()
        self.cmb_wait.addItems(["开始前等待", "末尾等待"])
        self.cmb_wait.setCurrentIndex(0)
        row_top.addWidget(self.cmb_wait)

        row_top.addStretch()

        # 步骤表：新增“起始需等区域ID(可选)”和“填充颜色(可选)”
        self.tbl = QTableWidget(0, 9, self)
        self.tbl.setHorizontalHeaderLabels([
            "序号", "工序显示名", "工位组名", "并行能力",
            "时长(秒，逗号分隔表示多台)", "区域ID(可选)", "区域容量(可选)",
            "起始需等区域ID(可选)", "填充颜色(可选)"
        ])
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setColumnWidth(self.COL_COLOR, 40)

        root.addWidget(self.tbl, 1)

        # 底部导出按钮栏
        btn_bar = QHBoxLayout()
        root.addLayout(btn_bar)
        btn_bar.addStretch()
        self.btn_export = QPushButton("生成并导出组合票")
        btn_bar.addWidget(self.btn_export)

        # 状态栏（用于显示导出进度 / 完成信息）
        self.status = QStatusBar()
        self.setStatusBar(self.status)

    def _connect_signals(self):
        self.act_back.triggered.connect(self.go_home)
        self.act_add_row.triggered.connect(self.add_row)
        self.act_del_row.triggered.connect(self.del_row)
        self.act_fill_sample.triggered.connect(self.fill_sample)
        self.btn_export.clicked.connect(self.do_export)
        self.act_help.triggered.connect(self.show_help)
        self.act_diagram.triggered.connect(self.show_diagram)

    # ------------- 动作 ------------- #
    def add_row(self):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)
        # 默认值：序号递增、能力=1、区域留空
        self.tbl.setItem(r, 0, QTableWidgetItem(str(r + 1)))
        self.tbl.setItem(r, 1, QTableWidgetItem(""))
        self.tbl.setItem(r, 2, QTableWidgetItem(""))
        self.tbl.setItem(r, 3, QTableWidgetItem("1"))
        self.tbl.setItem(r, 4, QTableWidgetItem(""))
        self.tbl.setItem(r, 5, QTableWidgetItem(""))   # 区域ID(可选)
        self.tbl.setItem(r, 6, QTableWidgetItem(""))   # 区域容量(可选)
        self.tbl.setItem(r, 7, QTableWidgetItem(""))   # 起始需等区域ID(可选)

        color_btn = QPushButton("…")
        color_btn.setFixedSize(30, 22)
        color_btn.clicked.connect(lambda _, row=r: self._choose_color(row))
        self.tbl.setCellWidget(r, self.COL_COLOR, color_btn)
        color_item = QTableWidgetItem("")
        color_item.setData(Qt.UserRole, "")
        self.tbl.setItem(r, self.COL_COLOR, color_item)

        # self.refresh_diagram()

    def _choose_color(self, row: int):
        dlg_col = QColorDialog.getColor(parent=self)
        if dlg_col.isValid():
            hex_code = dlg_col.name()
            btn = self.tbl.cellWidget(row, self.COL_COLOR)
            btn.setStyleSheet(f"background:{hex_code};")
            self.tbl.item(row, self.COL_COLOR).setData(Qt.UserRole, hex_code)
        # self.refresh_diagram()

    def del_row(self):
        r = self.tbl.currentRow()
        if r >= 0:
            self.tbl.removeRow(r)
        # self.refresh_diagram()

    def fill_sample(self):
        """
        串行示例 + 阻塞区域 + 上游闸门：
        - Z1: [电检2] ~ [NDA圈内] 属于同一区域，容量=1
        - L2++ / 电检准备：起始需等区域ID = Z1  （即：开工前要等 Z1 有空位）
        """
        self.tbl.setRowCount(0)
        sample_rows = [
            # 序号, 显示名,   组,       能力, 时长,   区域ID, 容量, 起始需等区域
            ("1",  "L2++",   "L2++",   "1", "112", "",   "",   "Z1"),
            ("2",  "电检准备", "电检准备", "1", "39.5", "",   "",   "Z1"),
            ("3",  "电检1",   "电检",   "1", "80",  "",   "",   ""),   # 普通工位
            ("4",  "电检2",   "电检",   "1", "70",  "Z1", "1", ""),   # 区域入口
            ("5",  "电检结束", "电检结束", "1", "29.5","Z1", "",  ""),   # 区域内
            ("6",  "NDA圈外", "NDA外",   "1", "30",  "Z1", "",  ""),   # 区域内
            ("7",  "NDA圈内", "NDA内",   "1", "20",  "Z1", "",  ""),   # 区域出口
            ("8",  "NDA检查", "NDA检查", "1", "30",  "",   "",   ""),   # 区域外
        ]
        for row in sample_rows:
            self.add_row()
            r = self.tbl.rowCount() - 1
            for c, text in enumerate(row):
                self.tbl.setItem(r, c, QTableWidgetItem(str(text)))

        if not self.ed_project.text().strip():
            self.ed_project.setText("L2++")
        self.spn_cars.setValue(4)
        self.cmb_grid.setCurrentText("1.0")
        self.cmb_wait.setCurrentText("开始前等待")
        # self.refresh_diagram()

    def _collect_inputs(self):
        project = self.ed_project.text().strip() or "工程"
        cars = int(self.spn_cars.value())
        try:
            grid_step = float(self.cmb_grid.currentText())
            if grid_step <= 0:
                grid_step = 1.0
        except Exception:
            grid_step = 1.0
        wait_policy = "before" if self.cmb_wait.currentIndex() == 0 else "after"

        defs = []
        for r in range(self.tbl.rowCount()):
            seq = (self.tbl.item(r, 0).text().strip() if self.tbl.item(r, 0) else "")
            name = (self.tbl.item(r, 1).text().strip() if self.tbl.item(r, 1) else "")
            grp  = (self.tbl.item(r, 2).text().strip() if self.tbl.item(r, 2) else "")
            cap  = (self.tbl.item(r, 3).text().strip() if self.tbl.item(r, 3) else "1")
            dur  = (self.tbl.item(r, 4).text().strip() if self.tbl.item(r, 4) else "")
            zid  = (self.tbl.item(r, 5).text().strip() if self.tbl.item(r, 5) else "")
            zcap = (self.tbl.item(r, 6).text().strip() if self.tbl.item(r, 6) else "")
            gzd  = (self.tbl.item(r, 7).text().strip() if self.tbl.item(r, 7) else "")
            color_hex = self.tbl.item(r, self.COL_COLOR).data(Qt.UserRole) or ""

            if not name or not grp or not dur:
                continue

            try:
                capacity = max(1, int(float(cap)))
            except Exception:
                capacity = 1

            durations = []
            for part in dur.replace("，", ",").split(","):
                t = part.strip()
                if not t:
                    continue
                try:
                    durations.append(float(t))
                except Exception:
                    pass
            if not durations:
                continue

            rec = {
                "seq": int(float(seq)) if seq else len(defs) + 1,
                "display": name,
                "group": grp,
                "capacity": capacity,
                "durations": durations,
                "color": color_hex,
            }

            if zid:
                rec["zone_id"] = zid
                try:
                    rec["zone_capacity"] = max(1, int(float(zcap))) if zcap else 1
                except Exception:
                    rec["zone_capacity"] = 1

            if gzd:
                rec["gate_zone_id"] = gzd

            defs.append(rec)

        defs.sort(key=lambda x: x["seq"])
        if not defs:
            raise ValueError("请至少填写一行有效的步骤（工序显示名/工位组名/时长）")

        return project, cars, grid_step, wait_policy, defs

    def do_export(self):
        try:
            project, cars, grid_step, wait_policy, defs = self._collect_inputs()
        except Exception as e:
            QMessageBox.warning(self, "输入有误", str(e))
            return

        path, _ = QFileDialog.getSaveFileName(self, "导出位置", f"{project}_组合票.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        self.dst_path = path

        worker = Worker(
            tickets.schedule_and_export,
            defs, cars, grid_step, wait_policy, project, self.dst_path
        )
        worker.signals.error.connect(self._on_error)
        worker.signals.finished.connect(self._on_export_finished)
        self.thread_pool.start(worker)
        self.status.showMessage("正在生成组合票...", 5000)

    def _on_export_finished(self):
        self.status.showMessage("导出完成", 6000)
        QMessageBox.information(self, "完成", f"已导出：\n{self.dst_path}")

    def go_home(self):
        home = getattr(self, "home_window", None)
        if home is not None and hasattr(home, "show"):
            try:
                home.show()
            except Exception:
                pass
        self.close()

    # ---------- 帮助弹窗 ----------
    def show_help(self):
        msg = (
            "<h3>组合票操作指南</h3>"
            "<ol>"
            "<li>点击『添加步骤』逐行录入；列含 ‘区域ID/容量’ 与 ‘起始需等区域ID’</li>"
            "<li>同一阻塞段填写同一区域ID，并仅在段首行写容量</li>"
            "<li>若需闸门等待，在 ‘起始需等区域ID’ 填下游区名</li>"
            "<li>最后一列『…』可自选颜色；未选自动配色</li>"
            "<li>顶部参数：车号数量 / 时间格刻度 / 等待分配方式</li>"
            "<li>填写完点击『生成并导出组合票』即可生成 Excel</li>"
            "</ol>"
        )
        QMessageBox.information(self, "帮助", msg)

    # ---------- 流程图弹窗 ----------
    def show_diagram(self):
        """弹出彩色流程图对话框"""
        steps = []
        for r in range(self.tbl.rowCount()):
            name = (self.tbl.item(r, 1).text().strip() if self.tbl.item(r, 1) else "")
            zone = (self.tbl.item(r, 5).text().strip() if self.tbl.item(r, 5) else "")
            gate = (self.tbl.item(r, 7).text().strip() if self.tbl.item(r, 7) else "")
            color = self.tbl.item(r, self.COL_COLOR).data(Qt.UserRole) or "#b0bec5"
            if name:
                steps.append({"row": r, "name": name, "zone": zone, "gate": gate, "color": color})
        if not steps:
            QMessageBox.information(self, "提示", "请先录入步骤")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("流程示意")
        dlg.resize(900, 260)
        lay = QVBoxLayout(dlg)

        view = QGraphicsView()
        view.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing)
        lay.addWidget(view)

        scene = QGraphicsScene(view)
        self._draw_blocks(scene, steps)
        view.setScene(scene)
        view.fitInView(scene.sceneRect(), Qt.KeepAspectRatio)

        dlg.exec()

    # ---------- 串联示意图 ---------- #
    def _draw_blocks(self, scene: QGraphicsScene, steps):
        x, y = 0.0, 0.0
        h, w, gap = 32, 110, 22
        pen = QPen(Qt.black, 1)
        zone_cache = {}
        for idx, s in enumerate(steps):
            # 决定颜色：自选 > 同区
            col = s["color"] if s["color"] else zone_cache.get(s["zone"], "#90a4ae")
            if s["zone"] and s["zone"] not in zone_cache:
                zone_cache[s["zone"]] = col
            rect = QGraphicsRectItem(x, y, w, h)
            rect.setBrush(QBrush(QColor(col)))
            rect.setPen(pen)
            rect.setData(0, s["row"])
            rect.setFlag(QGraphicsRectItem.ItemIsSelectable, True)
            scene.addItem(rect)

            txt = QGraphicsTextItem(s["name"])
            txt.setDefaultTextColor(Qt.black)
            txt.setPos(x + 4, y + 4)
            scene.addItem(txt)

            if s["gate"]:
                gate_txt = QGraphicsTextItem("⛔")
                gate_txt.setDefaultTextColor(Qt.red)
                gate_txt.setPos(x - 14, y + 4)
                scene.addItem(gate_txt)

            if idx < len(steps) - 1:
                line = QGraphicsLineItem(x + w, y + h / 2, x + w + gap, y + h / 2)
                line.setPen(pen)
                scene.addItem(line)
                # 箭头
                path = QPainterPath()
                path.moveTo(QPointF(x + w + gap, y + h / 2))
                path.lineTo(QPointF(x + w + gap - 6, y + h / 2 - 4))
                path.lineTo(QPointF(x + w + gap - 6, y + h / 2 + 4))
                path.closeSubpath()
                scene.addPath(path, pen, QBrush(Qt.black))

            # 点击彩块自动选中行
            def make_cb(row_idx):
                return lambda _: self.tbl.selectRow(row_idx)
            rect.mousePressEvent = make_cb(s["row"])

            x += w + gap

    # def refresh_diagram(self):
    #     """根据表格内容重绘右侧流程示意"""
    #     steps = []
    #     for r in range(self.tbl.rowCount()):
    #         name_item = self.tbl.item(r, 1)
    #         zone_item = self.tbl.item(r, 5)
    #         gate_item = self.tbl.item(r, 7)
    #         name = name_item.text().strip() if name_item else ""
    #         zone = zone_item.text().strip() if zone_item else ""
    #         gate = gate_item.text().strip() if gate_item else ""
    #         color_hex = self.tbl.item(r, self.COL_COLOR).data(Qt.UserRole) or "#b0bec5"
    #         if name:
    #             steps.append({"row": r, "name": name, "zone": zone, "gate": gate, "color": color_hex})
    #     scene = QGraphicsScene(self.view)
    #     self._draw_blocks(scene, steps)
    #     self.view.setScene(scene)
    #     self.view.fitInView(scene.sceneRect(), Qt.KeepAspectRatio)

    def _on_error(self, tb: str):
        QMessageBox.critical(self, "出错了", tb)
        self.status.showMessage("发生错误", 6000)