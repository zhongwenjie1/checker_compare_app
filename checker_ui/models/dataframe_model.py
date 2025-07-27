from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PySide6.QtGui import QColor
import pandas as pd

class DataFrameModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame | None = None, status_col: str | None = None):
        super().__init__()
        self._df = df if df is not None else pd.DataFrame()
        self._status_col = status_col

    def setDataFrame(self, df: pd.DataFrame | None):
        self.beginResetModel()
        self._df = df if df is not None else pd.DataFrame()
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        if role == Qt.DisplayRole:
            v = self._df.iat[index.row(), index.column()]
            return "" if pd.isna(v) else str(v)

        if role == Qt.BackgroundRole and self._status_col and self._status_col in self._df.columns:
            status = self._df.iloc[index.row()][self._status_col]
            if status == "OK":
                return QColor("#E8F5E9")
            elif status == "NG":
                return QColor("#FFEBEE")
            elif status == "未比对":
                return QColor("#F5F5F5")

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            try:
                return str(self._df.columns[section])
            except Exception:
                return ""
        else:
            try:
                return str(self._df.index[section])
            except Exception:
                return ""