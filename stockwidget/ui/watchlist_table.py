"""自选表格的拖放源；行移动由 WatchlistEditor 统一完成。"""

from PySide6.QtCore import Qt
from PySide6.QtGui import QCursor, QDrag
from PySide6.QtWidgets import QTableWidget
from shiboken6 import isValid


class WatchlistTable(QTableWidget):
    def startDrag(self, supported_actions):
        indexes = self.selectedIndexes()
        if not indexes or not supported_actions & Qt.MoveAction:
            return
        mime_data = self.model().mimeData(indexes)
        if mime_data is None:
            return

        rect = self.visualRect(indexes[0])
        for index in indexes[1:]:
            rect = rect.united(self.visualRect(index))
        drag = QDrag(self)
        drag.setMimeData(mime_data)
        drag.setPixmap(self.viewport().grab(rect))
        drag.setHotSpot(self.viewport().mapFromGlobal(QCursor.pos()) - rect.topLeft())
        try:
            # 不调用 QAbstractItemView.startDrag：它在 exec 返回 MoveAction 后
            # 会再次删除选中行。macOS 的嵌套事件循环使定时恢复选中无法规避此行为。
            drag.exec(Qt.MoveAction, Qt.MoveAction)
        finally:
            if isValid(self):
                self.setState(QTableWidget.State.NoState)
                self.viewport().update()
