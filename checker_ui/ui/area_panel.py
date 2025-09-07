# checker_ui/ui/area_panel.py
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QLabel, QHeaderView, QMessageBox
)
from PySide6.QtCore import Qt
from ..models.state import AppState

class AreaPanel(QWidget):
    """区域→岗位映射 + 岗位节拍(s) 的侧栏表格"""
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state

        self.table = QTableWidget(0, 3, self)
        self.table.setHorizontalHeaderLabels(["区域ID", "岗位ID", "岗位节拍(s)"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked |
                                   QTableWidget.EditTrigger.SelectedClicked)

        btn_add = QPushButton("添加行")
        btn_del = QPushButton("删除选中")
        btn_save = QPushButton("保存")
        btn_add.clicked.connect(self._on_add)
        btn_del.clicked.connect(self._on_del)
        btn_save.clicked.connect(self._on_save)

        hint = QLabel("提示：区域ID会在导出前映射成岗位ID；若配置了岗位节拍，将在该岗位的区域入口限流（不改变作业时长）。")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#555;")

        lay = QVBoxLayout(self)
        lay.addWidget(hint)
        lay.addWidget(self.table, 1)
        bar = QHBoxLayout()
        bar.addWidget(btn_add)
        bar.addWidget(btn_del)
        bar.addStretch(1)
        bar.addWidget(btn_save)
        lay.addLayout(bar)

        self._load_to_table()

    def _load_to_table(self):
        ats = dict(self.state.area_to_station or {})
        sct = dict(self.state.station_cycle_times or {})
        all_areas = sorted(set(ats.keys())) or [""]
        self.table.setRowCount(len(all_areas))
        for r, area in enumerate(all_areas):
            station = ats.get(area, "")
            cycle = sct.get(station, None) if station else None
            self.table.setItem(r, 0, QTableWidgetItem(str(area)))
            self.table.setItem(r, 1, QTableWidgetItem(str(station)))
            self.table.setItem(r, 2, QTableWidgetItem("" if cycle in (None, "") else str(cycle)))

    def _on_add(self):
        r = self.table.rowCount()
        self.table.insertRow(r)
        for c in range(3):
            self.table.setItem(r, c, QTableWidgetItem(""))

    def _on_del(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)

    def _on_save(self):
        area_to_station = {}
        station_cycle_times = dict(self.state.station_cycle_times or {})
        n = self.table.rowCount()
        for r in range(n):
            a_item = self.table.item(r, 0)
            s_item = self.table.item(r, 1)
            c_item = self.table.item(r, 2)
            area = (a_item.text() if a_item else "").strip()
            station = (s_item.text() if s_item else "").strip()
            cycle_txt = (c_item.text() if c_item else "").strip()
            if not area and not station and not cycle_txt:
                continue
            if area and station:
                area_to_station[area] = station
            if station and cycle_txt:
                try:
                    val = float(cycle_txt)
                    if val > 0:
                        station_cycle_times[station] = val
                except Exception:
                    pass

        self.state.area_to_station = area_to_station
        self.state.station_cycle_times = station_cycle_times
        try:
            self.state.save_state()
            QMessageBox.information(self, "已保存", "区域/岗位映射与岗位节拍已保存。")
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"保存出错：{e}")