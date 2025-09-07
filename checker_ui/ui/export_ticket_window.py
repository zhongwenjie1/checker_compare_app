# -*- coding: utf-8 -*-
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QToolBar, QStatusBar, QMessageBox, QTableWidget,
    QTableWidgetItem, QSpinBox, QComboBox, QLineEdit, QColorDialog,
    QGraphicsScene, QGraphicsView, QGraphicsRectItem, QGraphicsTextItem, QGraphicsLineItem, QSplitter, QDialog,
    QDockWidget, QFormLayout, QDoubleSpinBox, QDialogButtonBox, QMenu, QAbstractItemView
)
from PySide6.QtCore import Qt, QThreadPool, QPointF

from PySide6.QtGui import QAction, QColor, QPainter, QPainterPath, QPen, QBrush

# 功能开关：是否显示右侧“区域面板”
FEATURE_SHOW_AREA_PANEL = False


try:
    from checker_ui.infra.threads import Worker
    from checker_ui.core import tickets
except Exception:
    from ..infra.threads import Worker
    from ..core import tickets


try:
    from checker_ui.ui.workstation_dialog import WorkstationDialog
except Exception:
    from ..ui.workstation_dialog import WorkstationDialog

# FlowView fallback import
try:
    from checker_ui.ui.flow_view import FlowView
except Exception:
    from ..ui.flow_view import FlowView


# 区域面板（仅用于本窗口表格）
class _ExportAreaPanel(QWidget):
    """右侧区域面板（仅作用于本窗口的表格）。
    - 列出当前表格中出现的所有 区域ID；
    - 为选中区域批量设置“节拍时间(秒)”（会写入该区域所有行的第6列）。
    """
    def __init__(self, host_window: 'ExportTicketWindow'):
        super().__init__(host_window)
        self.host = host_window
        self._programmatic = False  # 防止 itemChanged 的循环触发

        lay = QFormLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(8)

        self.cmb_zone = QComboBox(self)
        lay.addRow("区域ID：", self.cmb_zone)

        self.spn_cycle = QDoubleSpinBox(self)
        self.spn_cycle.setRange(0.0, 99999.0)
        self.spn_cycle.setDecimals(1)
        self.spn_cycle.setSingleStep(0.5)
        self.spn_cycle.setValue(0.0)
        lay.addRow("节拍时间(秒)：", self.spn_cycle)

        self.btn_apply = QPushButton("应用节拍到该区域", self)
        lay.addRow("", self.btn_apply)

        self.cmb_zone.currentTextChanged.connect(self._sync_cycle_from_table)
        self.btn_apply.clicked.connect(self._apply_cycle_to_zone)

    def rebuild_zones(self):
        zones = []
        seen = set()
        tbl = self.host.tbl
        for r in range(tbl.rowCount()):
            item = tbl.item(r, 6)  # 区域ID列
            zid = item.text().strip() if item else ""
            if zid and zid not in seen:
                seen.add(zid)
                zones.append(zid)
        self._programmatic = True
        try:
            self.cmb_zone.clear()
            self.cmb_zone.addItems(zones)
        finally:
            self._programmatic = False
        # 选区后同步一次节拍显示
        if zones:
            self._sync_cycle_from_table(zones[0])
        else:
            self.spn_cycle.setValue(0.0)

    def _sync_cycle_from_table(self, zid: str):
        if self._programmatic:
            return
        if not zid:
            self.spn_cycle.setValue(0.0)
            return
        # 读取该区域在表中的最大节拍（为空按0）
        tbl = self.host.tbl
        max_ct = 0.0
        for r in range(tbl.rowCount()):
            zitem = tbl.item(r, 6)
            if not zitem:
                continue
            if (zitem.text().strip() == zid):
                ct_item = tbl.item(r, 5)
                if ct_item:
                    try:
                        v = float(ct_item.text().strip())
                        if v > max_ct:
                            max_ct = v
                    except Exception:
                        pass
        self.spn_cycle.setValue(max_ct)

    def _apply_cycle_to_zone(self):
        zid = self.cmb_zone.currentText().strip()
        if not zid:
            return
        val = float(self.spn_cycle.value())
        tbl = self.host.tbl
        self._programmatic = True
        try:
            tbl.blockSignals(True)
            for r in range(tbl.rowCount()):
                zitem = tbl.item(r, 6)
                if not zitem:
                    continue
                if zitem.text().strip() == zid:
                    # 写入第6列：节拍时间(秒，可选)
                    item = tbl.item(r, 5)
                    if item is None:
                        item = QTableWidgetItem("")
                        tbl.setItem(r, 5, item)
                    # 统一保留1位小数（与导出网格0.5/1.0常见设置匹配）
                    item.setText(f"{val:.1f}" if abs(val - round(val)) > 1e-9 else str(int(round(val))))
        finally:
            tbl.blockSignals(False)
            self._programmatic = False
        # 应用后保持列表同步
        self._sync_cycle_from_table(zid)

    # 供表格变动时调用
    def on_table_changed(self, *_):
        if not self._programmatic:
            self.rebuild_zones()


# ========== 岗位属性对话框（编辑 & 预览） ==========
class StationPropertiesDialog(QDialog):
    """
    输入/输出 payload 结构兼容 _insert_workstation_rows：
      {
        station_id, station_display, zone_id, color, cycle_time, zone_capacity,
        steps: [{name, duration}, ...],
        gate: {zone_id}
      }
    """
    def __init__(self, parent, initial: dict):
        super().__init__(parent)
        self.setWindowTitle("岗位属性")
        self.resize(680, 520)
        self._color_hex = initial.get("color") or ""

        lay = QVBoxLayout(self)

        # 顶部：基础属性
        row = QHBoxLayout(); lay.addLayout(row)
        row.addWidget(QLabel("工位组名："))
        self.ed_station = QLineEdit(initial.get("station_id", "")); self.ed_station.setReadOnly(True); self.ed_station.setFixedWidth(150); row.addWidget(self.ed_station)
        row.addSpacing(10); row.addWidget(QLabel("显示名："))
        self.ed_display = QLineEdit(initial.get("station_display", "") or initial.get("station_id", "")); self.ed_display.setFixedWidth(180); row.addWidget(self.ed_display)
        row.addSpacing(10); row.addWidget(QLabel("区域ID："))
        self.ed_zone = QLineEdit(initial.get("zone_id", "")); self.ed_zone.setFixedWidth(120); row.addWidget(self.ed_zone)

        row2 = QHBoxLayout(); lay.addLayout(row2)
        row2.addWidget(QLabel("节拍(秒)："))
        self.spn_cycle = QDoubleSpinBox(); self.spn_cycle.setRange(0, 99999); self.spn_cycle.setDecimals(1); self.spn_cycle.setSingleStep(0.5)
        self.spn_cycle.setValue(float(initial.get("cycle_time") or 0)); self.spn_cycle.setFixedWidth(120); row2.addWidget(self.spn_cycle)

        row2.addSpacing(10); row2.addWidget(QLabel("区域容量："))
        self.spn_cap = QSpinBox(); self.spn_cap.setRange(1, 99); self.spn_cap.setValue(int(initial.get("zone_capacity") or 1)); self.spn_cap.setFixedWidth(120); row2.addWidget(self.spn_cap)

        row2.addSpacing(10); row2.addWidget(QLabel("起始需等区域ID："))
        self.ed_gate = QLineEdit((initial.get("gate") or {}).get("zone_id", "")); self.ed_gate.setFixedWidth(140); row2.addWidget(self.ed_gate)

        row2.addSpacing(10)
        self.btn_color = QPushButton("颜色…")
        if self._color_hex: self.btn_color.setStyleSheet(f"background:{self._color_hex};")
        self.btn_color.clicked.connect(self._choose_color); row2.addWidget(self.btn_color)
        row2.addStretch()

        # 中部：步骤表
        lay.addWidget(QLabel("步骤清单（名称 / 时长秒）"))
        self.tbl = QTableWidget(0, 2, self)
        self.tbl.setHorizontalHeaderLabels(["步骤名", "时长(秒)"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        lay.addWidget(self.tbl, 1)

        for st in (initial.get("steps") or []):
            r = self.tbl.rowCount()
            self.tbl.insertRow(r)
            self.tbl.setItem(r, 0, QTableWidgetItem(str(st.get("name", ""))))
            self.tbl.setItem(r, 1, QTableWidgetItem(str(st.get("duration", ""))))

        # 步骤按钮
        row3 = QHBoxLayout(); lay.addLayout(row3)
        self.btn_add = QPushButton("添加步骤"); self.btn_del = QPushButton("删除步骤"); self.btn_preview = QPushButton("预览")
        row3.addWidget(self.btn_add); row3.addWidget(self.btn_del); row3.addStretch(); row3.addWidget(self.btn_preview)
        self.btn_add.clicked.connect(self._add_row)
        self.btn_del.clicked.connect(self._del_row)
        self.btn_preview.clicked.connect(self._preview)

        # 按钮组
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _choose_color(self):
        col = QColorDialog.getColor(parent=self)
        if col.isValid():
            self._color_hex = col.name()
            self.btn_color.setStyleSheet(f"background:{self._color_hex};")

    def _add_row(self):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)
        self.tbl.setItem(r, 0, QTableWidgetItem(""))
        self.tbl.setItem(r, 1, QTableWidgetItem("0"))

    def _del_row(self):
        r = self.tbl.currentRow()
        if r >= 0:
            self.tbl.removeRow(r)

    def payload(self) -> dict:
        steps = []
        for r in range(self.tbl.rowCount()):
            name = self.tbl.item(r, 0).text().strip() if self.tbl.item(r, 0) else ""
            dur_txt = self.tbl.item(r, 1).text().strip() if self.tbl.item(r, 1) else "0"
            try: dur = float(dur_txt)
            except Exception: dur = 0.0
            steps.append({"name": name, "duration": dur})
        return {
            "station_id": self.ed_station.text().strip(),
            "station_display": self.ed_display.text().strip() or self.ed_station.text().strip(),
            "zone_id": self.ed_zone.text().strip(),
            "color": self._color_hex,
            "cycle_time": float(self.spn_cycle.value()),
            "zone_capacity": int(self.spn_cap.value()),
            "steps": steps,
            "gate": {"zone_id": self.ed_gate.text().strip()} if self.ed_gate.text().strip() else {},
        }

    def _preview(self):
        steps = self.payload().get("steps") or []
        if not steps:
            QMessageBox.information(self, "预览", "请先添加步骤"); return
        dlg = QDialog(self); dlg.setWindowTitle("预览"); dlg.resize(780, 240)
        v = QVBoxLayout(dlg)
        view = QGraphicsView(); view.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing)
        v.addWidget(view)
        scene = QGraphicsScene(view)
        x, y, w, h, gap = 0, 0, 110, 32, 22
        pen = QPen(Qt.black, 1)
        color = self._color_hex or "#90a4ae"
        for i, st in enumerate(steps):
            rect = QGraphicsRectItem(x, y, w, h); rect.setBrush(QBrush(QColor(color))); rect.setPen(pen); scene.addItem(rect)
            txt = QGraphicsTextItem(st.get("name") or ("(显示名)" if i == 0 else "")); txt.setDefaultTextColor(Qt.black); txt.setPos(x + 4, y + 4); scene.addItem(txt)
            if i < len(steps) - 1:
                line = QGraphicsLineItem(x + w, y + h/2, x + w + gap, y + h/2); line.setPen(pen); scene.addItem(line)
                arrow = QPainterPath(); arrow.moveTo(QPointF(x + w + gap, y + h/2))
                arrow.lineTo(QPointF(x + w + gap - 6, y + h/2 - 4)); arrow.lineTo(QPointF(x + w + gap - 6, y + h/2 + 4)); arrow.closeSubpath()
                scene.addPath(arrow, pen, QBrush(Qt.black))
            x += w + gap
        view.setScene(scene); view.fitInView(scene.sceneRect(), Qt.KeepAspectRatio)
        dlg.exec()


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
    COL_COLOR = 9

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("导出组合票")
        self.resize(1120, 720)

        self.thread_pool = QThreadPool.globalInstance()
        self.dst_path = None

        self._build_ui()
        self._connect_signals()
        self._col_cache = None

    # ---------------- UI ---------------- #
    def _build_ui(self):
        tb = QToolBar("Ticket")
        self.addToolBar(tb)

        # 仅保留：返回主页 / 流程图 / 新增岗位 / 属性
        self.act_back = QAction("返回主页", self)
        tb.addAction(self.act_back)
        tb.addSeparator()

        self.act_diagram = QAction("流程图", self)
        tb.addAction(self.act_diagram)
        tb.addSeparator()

        self.act_add_station = QAction("新增岗位", self)
        tb.addAction(self.act_add_station)
        tb.addSeparator()

        self.act_edit_station = QAction("属性", self)
        tb.addAction(self.act_edit_station)
        tb.addSeparator()

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

        # 步骤表：新增“节拍时间(秒，可选)”和“起始需等区域ID(可选)”和“填充颜色(可选)”
        self.tbl = QTableWidget(0, 10, self)
        self.tbl.setHorizontalHeaderLabels([
            "序号", "工序显示名", "工位组名", "并行能力",
            "时长(秒，逗号分隔表示多台)", "节拍时间(秒，可选)", "区域ID(可选)", "区域容量(可选)",
            "起始需等区域ID(可选)", "填充颜色(可选)"
        ])
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setColumnWidth(self.COL_COLOR, 40)
        # 只读预览，由「新增岗位 / 属性」统一入口编辑
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)

        self.tbl.setVisible(False)  # 数据源保留，但不显示
        self.view = QGraphicsView(self)
        self.view.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing)
        root.addWidget(self.view, 1)

        # 底部导出按钮栏
        btn_bar = QHBoxLayout()
        root.addLayout(btn_bar)
        btn_bar.addStretch()
        self.btn_export = QPushButton("生成并导出组合票")
        btn_bar.addWidget(self.btn_export)

        # 状态栏（用于显示导出进度 / 完成信息）
        self.status = QStatusBar()
        self.setStatusBar(self.status)

        # 右侧：区域面板 Dock（针对本窗口表格做批量“节拍时间”设置）
        if FEATURE_SHOW_AREA_PANEL:
            self._area_panel = _ExportAreaPanel(self)
            self.dock_area_panel = QDockWidget("区域面板", self)
            self.dock_area_panel.setObjectName("dock_area_panel_export")
            self.dock_area_panel.setWidget(self._area_panel)
            self.dock_area_panel.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
            self.addDockWidget(Qt.RightDockWidgetArea, self.dock_area_panel)
            # 初始构建一次区域列表
            self._area_panel.rebuild_zones()

        # 初始流程图渲染
        self.refresh_diagram()

    def _connect_signals(self):
        self.act_back.triggered.connect(self.go_home)
        self.btn_export.clicked.connect(self.do_export)

        self.act_diagram.triggered.connect(self.show_diagram)
        self.act_add_station.triggered.connect(self._open_add_workstation)
        self.act_edit_station.triggered.connect(self._open_edit_properties)

        # 双击或右键 -> 属性
        self.tbl.cellDoubleClicked.connect(lambda *_: self._open_edit_properties())
        self.tbl.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tbl.customContextMenuRequested.connect(self._show_table_menu)
        self.tbl.itemChanged.connect(lambda *_: self.refresh_diagram())
        # 若区域面板存在（目前隐藏），保持同步
        if hasattr(self, "_area_panel"):
            self.tbl.itemChanged.connect(self._area_panel.on_table_changed)
    def _show_table_menu(self, pos):
        menu = QMenu(self)
        act_prop = menu.addAction("属性")
        act_prop.triggered.connect(self._open_edit_properties)
        menu.exec(self.tbl.viewport().mapToGlobal(pos))


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
        self.tbl.setItem(r, 5, QTableWidgetItem(""))   # 节拍时间(秒，可选)
        self.tbl.setItem(r, 6, QTableWidgetItem(""))   # 区域ID(可选)
        self.tbl.setItem(r, 7, QTableWidgetItem(""))   # 区域容量(可选)
        self.tbl.setItem(r, 8, QTableWidgetItem(""))   # 起始需等区域ID(可选)

        color_btn = QPushButton("…")
        color_btn.setFixedSize(30, 22)
        color_btn.clicked.connect(lambda _, row=r: self._choose_color(row))
        self.tbl.setCellWidget(r, self.COL_COLOR, color_btn)
        color_item = QTableWidgetItem("")
        color_item.setData(Qt.UserRole, "")
        self.tbl.setItem(r, self.COL_COLOR, color_item)

        self.refresh_diagram()
        if hasattr(self, "_area_panel"):
            self._area_panel.rebuild_zones()

    def _choose_color(self, row: int):
        dlg_col = QColorDialog.getColor(parent=self)
        if dlg_col.isValid():
            hex_code = dlg_col.name()
            btn = self.tbl.cellWidget(row, self.COL_COLOR)
            btn.setStyleSheet(f"background:{hex_code};")
            self.tbl.item(row, self.COL_COLOR).setData(Qt.UserRole, hex_code)
        self.refresh_diagram()

    def _prepare_color_cell(self, row, color_hex=None):
        color_btn = QPushButton("…")
        color_btn.setFixedSize(30, 22)
        color_btn.clicked.connect(lambda _, r=row: self._choose_color(r))
        self.tbl.setCellWidget(row, self.COL_COLOR, color_btn)
        color_item = QTableWidgetItem("")
        color_item.setData(Qt.UserRole, color_hex or "")
        if color_hex:
            color_btn.setStyleSheet(f"background:{color_hex};")
        self.tbl.setItem(row, self.COL_COLOR, color_item)

    # ---------- 新增岗位向导 ----------
    def _open_add_workstation(self):
        # 收集现有区域ID供下拉提示
        zones = set()
        for r in range(self.tbl.rowCount()):
            item = self.tbl.item(r, self._col("区域ID") or 6)
            zid = item.text().strip() if item else ""
            if zid:
                zones.add(zid)
        try:
            dlg = WorkstationDialog(self, sorted(zones))
        except Exception:
            # 回退（相对导入已在顶部处理）
            dlg = WorkstationDialog(self, sorted(zones))
        dlg.acceptedWithData.connect(self._insert_workstation_rows)
        dlg.exec()

    def _insert_workstation_rows(self, payload: dict):
        """把“新增岗位”对话框数据，转为多行插入表格。"""
        station_id = payload.get("station_id", "").strip()
        display = payload.get("station_display", station_id)
        zone_id = payload.get("zone_id", "").strip()
        color_hex = payload.get("color") or ""
        cycle = payload.get("cycle_time", None)
        zone_cap = int(payload.get("zone_capacity") or 1)
        steps = payload.get("steps") or []
        gate = payload.get("gate")
        insert = payload.get("insert") or {"mode": "append", "index": -1}

        if not station_id or not zone_id or not steps:
            return

        table = self.tbl
        # 计算插入位置
        mode = insert.get("mode")
        if mode == "before_selected":
            row0 = table.currentRow()
            if row0 < 0:
                row0 = table.rowCount()
        elif mode == "before_index":
            try:
                idx = int(insert.get("index", 1))
            except Exception:
                idx = 1
            row0 = max(0, min(table.rowCount(), idx - 1))
        else:
            row0 = table.rowCount()

        # 实际插入
        for i, st in enumerate(steps):
            table.insertRow(row0 + i)
            # 准备颜色列
            self._prepare_color_cell(row0 + i, color_hex if color_hex else None)

            # 写入各列（用模糊列匹配，避免列顺序变化问题）
            self._set_cell(row0 + i, "序号", "")  # 稍后统一重排
            if i == 0:
                # 若第一步在对话框里填写了名称，则优先使用；否则退回到岗位“显示名”
                self._set_cell(row0 + i, "工序显示名", (st.get("name", "") or display))
            else:
                self._set_cell(row0 + i, "工序显示名", st.get("name", ""))
            self._set_cell(row0 + i, "工位组名", station_id)
            self._set_cell(row0 + i, "并行能力", str(zone_cap))

            dur = st.get("duration", 0)
            self._set_cell(row0 + i, "时长", str(dur))

            if cycle not in (None, "", 0, 0.0):
                self._set_cell(row0 + i, "节拍", str(cycle))

            self._set_cell(row0 + i, "区域ID", zone_id)
            self._set_cell(row0 + i, "区域容量", str(zone_cap))

            if gate and gate.get("zone_id"):
                self._set_cell(row0 + i, "起始需等区域ID", str(gate.get("zone_id")))
                if gate.get("buffer"):
                    # 目前界面没有单独列记录缓冲，保持在 tickets 里生效即可
                    pass

        # 统一重排“序号”列并选中首行
        self._renumber_seq()
        table.setCurrentCell(row0, self._col("工序显示名") or 1)
        # 同步右侧区域面板
        if hasattr(self, "_area_panel"):
            self._area_panel.rebuild_zones()
        # 刷新流程图
        self.refresh_diagram()

    # ==== 列定位/取写工具 ====
    def _header_text(self, col: int) -> str:
        item = self.tbl.horizontalHeaderItem(col)
        return item.text() if item else ""

    def _col(self, key_substr: str):
        # 缓存首轮
        if self._col_cache is None:
            self._col_cache = {}
            for c in range(self.tbl.columnCount()):
                self._col_cache[self._header_text(c)] = c
        for text, c in self._col_cache.items():
            if key_substr in text:
                return c
        return None

    def _set_cell(self, row: int, key_substr: str, val: str):
        col = self._col(key_substr)
        if col is None:
            # 容错：若没找到，放弃本字段
            return
        item = self.tbl.item(row, col)
        if item is None:
            item = QTableWidgetItem()
        item.setText(val)
        self.tbl.setItem(row, col, item)

    def _get_cell_text(self, row: int, key_substr: str) -> str:
        col = self._col(key_substr)
        if col is None:
            return ""
        it = self.tbl.item(row, col)
        return it.text().strip() if it else ""

    def _renumber_seq(self):
        col = self._col("序号")
        if col is None:
            return
        for r in range(self.tbl.rowCount()):
            self.tbl.setItem(r, col, QTableWidgetItem(str(r + 1)))

    def del_row(self):
        r = self.tbl.currentRow()
        if r >= 0:
            self.tbl.removeRow(r)
        self.refresh_diagram()
        if hasattr(self, "_area_panel"):
            self._area_panel.rebuild_zones()

    def fill_sample(self):
        """
        串行示例 + 阻塞区域 + 上游闸门：
        - Z1: [电检2] ~ [NDA圈内] 属于同一区域，容量=1
        - L2++ / 电检准备：起始需等区域ID = Z1  （即：开工前要等 Z1 有空位）
        """
        self.tbl.setRowCount(0)
        sample_rows = [
            # 序号, 显示名,   组,       能力, 时长, 节拍时间, 区域ID, 容量, 起始需等区域
            ("1",  "L2++",   "L2++",   "1", "112", "", "",   "",   "Z1"),
            ("2",  "电检准备", "电检准备", "1", "39.5", "", "",   "",   "Z1"),
            ("3",  "电检1",   "电检",   "1", "80",  "", "",   "",   ""),   # 普通工位
            ("4",  "电检2",   "电检",   "1", "70",  "", "Z1", "1", ""),   # 区域入口
            ("5",  "电检结束", "电检结束", "1", "29.5","", "Z1", "",  ""),   # 区域内
            ("6",  "NDA圈外", "NDA外",   "1", "30",  "", "Z1", "",  ""),   # 区域内
            ("7",  "NDA圈内", "NDA内",   "1", "20",  "", "Z1", "",  ""),   # 区域出口
            ("8",  "NDA检查", "NDA检查", "1", "30",  "", "",   "",   ""),   # 区域外
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
        self.refresh_diagram()
        if hasattr(self, "_area_panel"):
            self._area_panel.rebuild_zones()

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
            cycle_time = (self.tbl.item(r, 5).text().strip() if self.tbl.item(r, 5) else "")
            zid  = (self.tbl.item(r, 6).text().strip() if self.tbl.item(r, 6) else "")
            zcap = (self.tbl.item(r, 7).text().strip() if self.tbl.item(r, 7) else "")
            gzd  = (self.tbl.item(r, 8).text().strip() if self.tbl.item(r, 8) else "")
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

            if cycle_time:
                try:
                    rec["cycle_time"] = float(cycle_time)
                except Exception:
                    pass

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
        """弹出流程图：用 FlowView 以“框+箭头”展示，并支持双击打开属性。"""
        defs = self._collect_defs_for_flow()
        if not defs:
            QMessageBox.information(self, "提示", "请先录入步骤")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("流程图")
        dlg.resize(960, 420)
        lay = QVBoxLayout(dlg)

        view = FlowView(dlg)
        lay.addWidget(view)
        view.render(defs)

        # 在流程图里双击岗位框 -> 打开对应组的属性编辑
        view.editRequested.connect(self._open_properties_for_group)

        dlg.exec()
    def _collect_defs_for_flow(self):
        """把表格转换为 FlowView 所需的最小字段列表。"""
        defs = []
        for r in range(self.tbl.rowCount()):
            name = self._get_cell_text(r, "工序显示名")
            grp  = self._get_cell_text(r, "工位组名") or name
            dur_txt = self._get_cell_text(r, "时长")
            color_hex = ""
            item_col = self.tbl.item(r, self.COL_COLOR)
            if item_col is not None:
                color_hex = item_col.data(Qt.UserRole) or ""

            if not name or not dur_txt:
                continue
            # 取第一个时长作为可视化用
            try:
                first = str(dur_txt).replace("，", ",").split(",")[0].strip()
                dur = float(first)
            except Exception:
                dur = 0.0
            defs.append({
                "display": name,
                "group": grp,
                "durations": [dur],
                "color": color_hex,
            })
        return defs

    def _open_properties_for_group(self, group_id: str):
        """根据岗位组名在表格中定位并打开属性编辑。"""
        for r in range(self.tbl.rowCount()):
            if self._get_cell_text(r, "工位组名") == group_id:
                self.tbl.selectRow(r)
                self._open_edit_properties()
                return
        QMessageBox.information(self, "提示", f"未找到组：{group_id}")

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

            # 并行能力：右上角 ×N 角标（cap > 1 显示）
            cap_val = max(1, int(s.get("cap", 1) or 1))
            if cap_val > 1:
                badge = QGraphicsTextItem(f"×{cap_val}")
                badge.setDefaultTextColor(Qt.black)
                badge.setPos(x + w - 26, y - 10)
                scene.addItem(badge)

            # 标题
            txt = QGraphicsTextItem(s["name"])
            txt.setDefaultTextColor(Qt.black)
            txt.setPos(x + 4, y + 4)
            scene.addItem(txt)

            # 区域ID：右下角小标签（如 Z1）
            if s.get("zone"):
                ztag = QGraphicsTextItem(f"{s['zone']}")
                ztag.setDefaultTextColor(QColor("#333333"))
                ztag.setPos(x + w - 28, y + h - 18)
                scene.addItem(ztag)

            # 闸门：显式→红色；自动推断→灰色
            gate_zone = s.get("gate") or ""
            gate_color = Qt.red
            if not gate_zone:
                gate_zone = s.get("gate_auto") or ""
                gate_color = QColor("#888888")
            if gate_zone:
                gate_txt = QGraphicsTextItem(f"⛔{gate_zone}")
                gate_txt.setDefaultTextColor(gate_color)
                gate_txt.setPos(x - 28, y - 10)   # 左上角外侧
                scene.addItem(gate_txt)

            # 连线 + 箭头
            if idx < len(steps) - 1:
                line = QGraphicsLineItem(x + w, y + h / 2, x + w + gap, y + h / 2)
                line.setPen(pen)
                scene.addItem(line)
                path = QPainterPath()
                path.moveTo(QPointF(x + w + gap, y + h / 2))
                path.lineTo(QPointF(x + w + gap - 6, y + h / 2 - 4))
                path.lineTo(QPointF(x + w + gap - 6, y + h / 2 + 4))
                path.closeSubpath()
                scene.addPath(path, pen, QBrush(Qt.black))

            # 点击选择行 / 双击打开属性
            def make_cb(row_idx):
                return lambda _: self.tbl.selectRow(row_idx)
            rect.mousePressEvent = make_cb(s["row"])

            def make_dbl_cb(row_idx):
                def _dbl(ev):
                    self.tbl.selectRow(row_idx)
                    self._open_edit_properties()
                return _dbl
            rect.mouseDoubleClickEvent = make_dbl_cb(s["row"])

            x += w + gap

    def refresh_diagram(self):
        """根据表格内容重绘中央流程图"""
        steps = []
        for r in range(self.tbl.rowCount()):
            name_item = self.tbl.item(r, 1)   # 工序显示名
            zone_item = self.tbl.item(r, 6)   # 区域ID(可选)
            gate_item = self.tbl.item(r, 8)   # 起始需等区域ID(可选)
            name = name_item.text().strip() if name_item else ""
            zone = zone_item.text().strip() if zone_item else ""
            gate = gate_item.text().strip() if gate_item else ""
            color_hex = ""
            color_it = self.tbl.item(r, self.COL_COLOR)
            if color_it is not None:
                color_hex = color_it.data(Qt.UserRole) or ""
            # 并行能力（默认1）
            cap_val = 1
            cap_item = self.tbl.item(r, 3)  # “并行能力”列
            try:
                cap_val = int(float(cap_item.text().strip())) if cap_item and cap_item.text().strip() else 1
                if cap_val < 1:
                    cap_val = 1
            except Exception:
                cap_val = 1

            if name:
                steps.append({
                    "row": r,
                    "name": name,
                    "zone": zone,
                    "gate": gate,
                    "cap": cap_val,
                    "color": color_hex or "#b0bec5",
                })
        # 如果最后一个结点是“结束”且它并非来自表格（防御性处理），去掉它
        if steps and steps[-1]["name"] == "结束":
            real_exist = False
            for r in range(self.tbl.rowCount()):
                if self._get_cell_text(r, "工序显示名") == "结束":
                    real_exist = True
                    break
            if not real_exist:
                steps.pop()

        # —— 自动判定：建议闸门 gate_auto（灰色提示）——
        # 规则：对“自己无 zone 且未显式 gate”的步骤，寻找**后续的第一个区域入口**，作为 gate_auto。
        def is_zone_entry(lst, idx):
            return lst[idx]["zone"] and (idx == 0 or lst[idx-1]["zone"] != lst[idx]["zone"])

        for i in range(len(steps)):
            steps[i]["gate_auto"] = ""
        for i in range(len(steps)):
            if steps[i]["zone"] or steps[i]["gate"]:
                continue  # 已在区域内或已显式闸门，不再自动提示
            for j in range(i + 1, len(steps)):
                if is_zone_entry(steps, j):
                    steps[i]["gate_auto"] = steps[j]["zone"]
                    break

        scene = QGraphicsScene(self.view)
        self._draw_blocks(scene, steps)
        self.view.setScene(scene)
        self.view.fitInView(scene.sceneRect(), Qt.KeepAspectRatio)

    def _on_error(self, tb: str):
        QMessageBox.critical(self, "出错了", tb)
        self.status.showMessage("发生错误", 6000)
    # ===== 岗位属性编辑 =====
    def _open_edit_properties(self):
        row = self.tbl.currentRow()
        if row < 0:
            QMessageBox.information(self, "提示", "请先选中任意一个步骤所在行，再点击『属性』。")
            return

        payload, block_rows = self._collect_station_payload_at(row)
        if not payload or not block_rows:
            QMessageBox.warning(self, "提示", "未能识别当前行所属的岗位块。请确认该行已填写『工位组名』。")
            return

        dlg = StationPropertiesDialog(self, payload)
        if dlg.exec() != QDialog.Accepted:
            return

        new_payload = dlg.payload()
        insert_before_index = block_rows[0] + 1   # 保留原位置
        new_payload["insert"] = {"mode": "before_index", "index": insert_before_index}

        # 删除旧块（从下往上删）
        for r in sorted(block_rows, reverse=True):
            self.tbl.removeRow(r)

        # 复用“新增岗位”插入逻辑
        self._insert_workstation_rows(new_payload)
        self._renumber_seq()
        # 修改后刷新流程图
        self.refresh_diagram()

    def _collect_station_payload_at(self, any_row: int):
        """
        以 any_row 为中心，按『工位组名』向上/向下扩展，取得连续块，组装 payload。
        返回 (payload: dict, block_rows: [int...])
        """
        col_grp = self._col("工位组名")
        if col_grp is None:
            return None, None
        grp = self._get_cell_text(any_row, "工位组名")
        if not grp:
            return None, None

        # 向上/下扩展
        r0 = any_row
        while r0 - 1 >= 0 and self._get_cell_text(r0 - 1, "工位组名") == grp:
            r0 -= 1
        r1 = any_row
        while r1 + 1 < self.tbl.rowCount() and self._get_cell_text(r1 + 1, "工位组名") == grp:
            r1 += 1
        rows = list(range(r0, r1 + 1))
        if not rows:
            return None, None

        # 公共属性（以首行为准）
        display = self._get_cell_text(rows[0], "工序显示名") or grp
        zone_id = self._get_cell_text(rows[0], "区域ID")
        cycle_txt = self._get_cell_text(rows[0], "节拍时间")
        gate_zone = self._get_cell_text(rows[0], "起始需等区域ID")
        cap_txt = self._get_cell_text(rows[0], "区域容量") or self._get_cell_text(rows[0], "并行能力")

        try: cap = max(1, int(float(cap_txt))) if cap_txt else 1
        except Exception: cap = 1
        try: cycle = float(cycle_txt) if cycle_txt else 0.0
        except Exception: cycle = 0.0

        color_hex = self.tbl.item(rows[0], self.COL_COLOR).data(Qt.UserRole) or ""

        # 组装 steps（与现有行数对齐；第 1 行 name 可留空）
        steps = []
        for i, r in enumerate(rows):
            name = self._get_cell_text(r, "工序显示名")
            dur_txt = self._get_cell_text(r, "时长")
            dur_val = 0.0
            if dur_txt:
                part = str(dur_txt).replace("，", ",").split(",")[0].strip()
                try: dur_val = float(part)
                except Exception: dur_val = 0.0
            steps.append({"name": name, "duration": dur_val})

        payload = {
            "station_id": grp,
            "station_display": display,
            "zone_id": zone_id,
            "color": color_hex,
            "cycle_time": cycle,
            "zone_capacity": cap,
            "steps": steps,
            "gate": {"zone_id": gate_zone} if gate_zone else {},
        }
        return payload, rows