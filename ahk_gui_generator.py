"""
AHK Script Builder (PySide6 + pyautogui)

Features:
- GUI to create named actions (Click, Sleep, Run, KeyPress, WaitWindow)
- Capture mouse coordinates for Click actions
- Save/load project (JSON)
- Export .ahk file that contains both actionable AHK commands and structured comment lines
- Load an existing .ahk file if it was created by this tool (parses comment ACTION lines)

Dependencies:
- PySide6
- pyautogui

Install:
    pip install PySide6 pyautogui

Run:
    python ahk_gui_generator.py

Notes:
- The exported .ahk will contain comment lines like:
    ;ACTION:Click;Name=Start;X=100;Y=200
  which the GUI will parse when loading such an .ahk file.

- This is a single-file starter app. You can extend action types, add image-search,
  or add Windows shortcut creation using pywin32 if desired.

"""
import sys
import json
import os
import time
from enum import Enum
from dataclasses import dataclass, asdict

from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QListWidget, QMessageBox, QFileDialog,
    QDialog, QFormLayout, QComboBox, QSpinBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QAction, QToolBar
)
from PySide6.QtCore import Qt

import pyautogui

APP_DIR = os.path.abspath(os.path.dirname(__file__))


class ActionType(str, Enum):
    CLICK = "Click"
    SLEEP = "Sleep"
    RUN = "Run"
    KEYPRESS = "KeyPress"
    WAITWINDOW = "WaitWindow"


@dataclass
class Action:
    type: ActionType
    name: str
    x: int = 0
    y: int = 0
    ms: int = 0
    command: str = ""

    def to_comment(self) -> str:
        # Structured comment line used by the GUI to reload
        if self.type == ActionType.CLICK:
            return f";ACTION:Click;Name={self.name};X={self.x};Y={self.y}"
        if self.type == ActionType.SLEEP:
            return f";ACTION:Sleep;Name={self.name};MS={self.ms}"
        if self.type == ActionType.RUN:
            return f";ACTION:Run;Name={self.name};Cmd={self.command}"
        if self.type == ActionType.KEYPRESS:
            return f";ACTION:KeyPress;Name={self.name};Key={self.command}"
        if self.type == ActionType.WAITWINDOW:
            return f";ACTION:WaitWindow;Name={self.name};Title={self.command}"
        return f";ACTION:Unknown;Name={self.name}"

    def to_ahk(self) -> str:
        # Real AHK commands implementing the action
        if self.type == ActionType.CLICK:
            # Using CoordMode, Mouse, Screen for absolute coords
            return f"MouseMove, {self.x}, {self.y}\nClick\n"
        if self.type == ActionType.SLEEP:
            return f"Sleep, {self.ms}\n"
        if self.type == ActionType.RUN:
            # Use Run
            return f"Run, {self.command}\n"
        if self.type == ActionType.KEYPRESS:
            return f"Send, {self.command}\n"
        if self.type == ActionType.WAITWINDOW:
            # Wait for window title
            return f"WinWait, {self.command}, , 10\nWinActivate, {self.command}\n"
        return "\n"


class ActionDialog(QDialog):
    def __init__(self, parent=None, action: Action | None = None):
        super().__init__(parent)
        self.setWindowTitle("Add / Edit Action")
        self.resize(400, 200)
        self.action = action

        self.layout = QFormLayout()

        self.type_cb = QComboBox()
        for t in ActionType:
            self.type_cb.addItem(t.value)
        self.layout.addRow("Type:", self.type_cb)

        self.name_edit = QLineEdit()
        self.layout.addRow("Name:", self.name_edit)

        # For click: coords
        self.x_spin = QSpinBox()
        self.x_spin.setRange(0, 99999)
        self.y_spin = QSpinBox()
        self.y_spin.setRange(0, 99999)
        xy_layout = QHBoxLayout()
        xy_layout.addWidget(QLabel("X:"))
        xy_layout.addWidget(self.x_spin)
        xy_layout.addWidget(QLabel("Y:"))
        xy_layout.addWidget(self.y_spin)
        self.capture_btn = QPushButton("Capture Mouse Position")
        self.capture_btn.clicked.connect(self.capture_position)
        xy_layout.addWidget(self.capture_btn)
        self.layout.addRow("Coordinates:", xy_layout)

        # Sleep
        self.ms_spin = QSpinBox()
        self.ms_spin.setRange(0, 6000000)
        self.ms_spin.setSingleStep(100)
        self.layout.addRow("Milliseconds:", self.ms_spin)

        # Command
        self.cmd_edit = QLineEdit()
        self.layout.addRow("Command/Key/Window Title:", self.cmd_edit)

        # Buttons
        btn_layout = QHBoxLayout()
        self.ok_btn = QPushButton("OK")
        self.cancel_btn = QPushButton("Cancel")
        self.ok_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.ok_btn)
        btn_layout.addWidget(self.cancel_btn)

        self.layout.addRow(btn_layout)
        self.setLayout(self.layout)

        self.type_cb.currentTextChanged.connect(self.type_changed)
        if action:
            self.load_action(action)
        else:
            self.type_changed(self.type_cb.currentText())

    def load_action(self, action: Action):
        self.type_cb.setCurrentText(action.type.value)
        self.name_edit.setText(action.name)
        self.x_spin.setValue(action.x)
        self.y_spin.setValue(action.y)
        self.ms_spin.setValue(action.ms)
        self.cmd_edit.setText(action.command)
        self.type_changed(action.type.value)

    def type_changed(self, text):
        # Show/hide fields based on type
        t = ActionType(text)
        self.x_spin.setEnabled(t == ActionType.CLICK)
        self.y_spin.setEnabled(t == ActionType.CLICK)
        self.capture_btn.setEnabled(t == ActionType.CLICK)
        self.ms_spin.setEnabled(t == ActionType.SLEEP)
        self.cmd_edit.setEnabled(t in (ActionType.RUN, ActionType.KEYPRESS, ActionType.WAITWINDOW))

    def capture_position(self):
        # Hide dialog, instruct user to move mouse and press OK to capture
        QMessageBox.information(self, "Capture position",
                                "After you click OK, move the mouse to the target position and press OK again to capture the coordinates.")
        # First OK to close the info box, then capture after user moves and presses OK
        res = QMessageBox.question(self, "Ready to capture", "Move mouse to target position and press OK to capture current mouse position.", QMessageBox.Ok | QMessageBox.Cancel)
        if res == QMessageBox.Ok:
            time.sleep(0.2)
            x, y = pyautogui.position()
            # pyautogui returns screen coords; we store those
            self.x_spin.setValue(x)
            self.y_spin.setValue(y)
            QMessageBox.information(self, "Captured", f"Captured position: X={x}, Y={y}")

    def get_action(self) -> Action:
        t = ActionType(self.type_cb.currentText())
        return Action(
            type=t,
            name=self.name_edit.text() or t.value,
            x=self.x_spin.value(),
            y=self.y_spin.value(),
            ms=self.ms_spin.value(),
            command=self.cmd_edit.text()
        )


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AHK Script Builder")
        self.resize(800, 500)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout()
        central.setLayout(layout)

        top_layout = QHBoxLayout()
        layout.addLayout(top_layout)

        top_layout.addWidget(QLabel("Application name (for comments / file):"))
        self.app_name_edit = QLineEdit("MyScript")
        top_layout.addWidget(self.app_name_edit)

        self.new_btn = QPushButton("New Action")
        self.new_btn.clicked.connect(self.add_action)
        top_layout.addWidget(self.new_btn)

        self.edit_btn = QPushButton("Edit")
        self.edit_btn.clicked.connect(self.edit_action)
        top_layout.addWidget(self.edit_btn)

        self.remove_btn = QPushButton("Remove")
        self.remove_btn.clicked.connect(self.remove_action)
        top_layout.addWidget(self.remove_btn)

        # Actions table
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["#", "Type", "Name", "Params", "Preview"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()
        layout.addLayout(btn_layout)

        self.load_project_btn = QPushButton("Load .ahk or project")
        self.load_project_btn.clicked.connect(self.load_file)
        btn_layout.addWidget(self.load_project_btn)

        self.save_project_btn = QPushButton("Save project (.json)")
        self.save_project_btn.clicked.connect(self.save_project)
        btn_layout.addWidget(self.save_project_btn)

        self.export_btn = QPushButton("Export .ahk")
        self.export_btn.clicked.connect(self.export_ahk)
        btn_layout.addWidget(self.export_btn)

        # Internal actions list
        self.actions: list[Action] = []

        # Toolbar actions
        toolbar = QToolBar()
        self.addToolBar(toolbar)
        new_act = QAction("New", self)
        new_act.triggered.connect(self.add_action)
        toolbar.addAction(new_act)

        open_act = QAction("Open", self)
        open_act.triggered.connect(self.load_file)
        toolbar.addAction(open_act)

        save_act = QAction("Save", self)
        save_act.triggered.connect(self.save_project)
        toolbar.addAction(save_act)

        export_act = QAction("Export .ahk", self)
        export_act.triggered.connect(self.export_ahk)
        toolbar.addAction(export_act)

    def add_action(self):
        dlg = ActionDialog(self)
        if dlg.exec() == QDialog.Accepted:
            action = dlg.get_action()
            self.actions.append(action)
            self.refresh_table()

    def edit_action(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.actions):
            QMessageBox.warning(self, "Select action", "Please select an action to edit.")
            return
        action = self.actions[row]
        dlg = ActionDialog(self, action)
        if dlg.exec() == QDialog.Accepted:
            self.actions[row] = dlg.get_action()
            self.refresh_table()

    def remove_action(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.actions):
            QMessageBox.warning(self, "Select action", "Please select an action to remove.")
            return
        del self.actions[row]
        self.refresh_table()

    def refresh_table(self):
        self.table.setRowCount(len(self.actions))
        for i, a in enumerate(self.actions):
            self.table.setItem(i, 0, QTableWidgetItem(str(i+1)))
            self.table.setItem(i, 1, QTableWidgetItem(a.type.value))
            self.table.setItem(i, 2, QTableWidgetItem(a.name))
            params = ""
            if a.type == ActionType.CLICK:
                params = f"X={a.x}, Y={a.y}"
            elif a.type == ActionType.SLEEP:
                params = f"MS={a.ms}"
            elif a.type in (ActionType.RUN, ActionType.KEYPRESS, ActionType.WAITWINDOW):
                params = a.command
            self.table.setItem(i, 3, QTableWidgetItem(params))
            self.table.setItem(i, 4, QTableWidgetItem(a.to_comment()))

    def save_project(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save project", os.path.join(APP_DIR, "project.json"), "JSON Files (*.json)")
        if not path:
            return
        data = {
            'app_name': self.app_name_edit.text(),
            'actions': [asdict(a) for a in self.actions]
        }
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        QMessageBox.information(self, "Saved", f"Project saved to {path}")

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open .ahk or project", APP_DIR, "AHK Scripts (*.ahk);;JSON Files (*.json);;All Files (*)")
        if not path:
            return
        if path.lower().endswith('.json'):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.app_name_edit.setText(data.get('app_name', 'MyScript'))
            self.actions = []
            for ad in data.get('actions', []):
                a = Action(type=ActionType(ad['type']), name=ad['name'], x=ad.get('x',0), y=ad.get('y',0), ms=ad.get('ms',0), command=ad.get('command',''))
                self.actions.append(a)
            self.refresh_table()
            QMessageBox.information(self, "Loaded", f"Project loaded from {path}")
            return

        # If .ahk, try to parse structured comment lines
        if path.lower().endswith('.ahk'):
            actions = []
            app_name = os.path.splitext(os.path.basename(path))[0]
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    line = line.strip()
                    if line.startswith(';ACTION:'):
                        # parse kv pairs separated by ;
                        try:
                            rest = line[len(';ACTION:'):]
                            parts = rest.split(';')
                            typ = parts[0]
                            kv = {}
                            for p in parts[1:]:
                                if '=' in p:
                                    k,v = p.split('=',1)
                                    kv[k]=v
                            if typ == 'Click':
                                a = Action(ActionType.CLICK, kv.get('Name','Click'), int(kv.get('X',0)), int(kv.get('Y',0)))
                            elif typ == 'Sleep':
                                a = Action(ActionType.SLEEP, kv.get('Name','Sleep'), ms=int(kv.get('MS',0)))
                            elif typ == 'Run':
                                a = Action(ActionType.RUN, kv.get('Name','Run'), command=kv.get('Cmd',''))
                            elif typ == 'KeyPress':
                                a = Action(ActionType.KEYPRESS, kv.get('Name','KeyPress'), command=kv.get('Key',''))
                            elif typ == 'WaitWindow':
                                a = Action(ActionType.WAITWINDOW, kv.get('Name','WaitWindow'), command=kv.get('Title',''))
                            else:
                                continue
                            actions.append(a)
                        except Exception as e:
                            print('Failed to parse ACTION line:', e)
            self.actions = actions
            self.app_name_edit.setText(app_name)
            self.refresh_table()
            QMessageBox.information(self, "Loaded .ahk", f"Loaded actions from {path} (found {len(actions)} actions)")
            return

        QMessageBox.warning(self, "Unsupported", "Unsupported file type")

    def export_ahk(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export .ahk", os.path.join(APP_DIR, f"{self.app_name_edit.text()}.ahk"), "AHK Scripts (*.ahk)")
        if not path:
            return
        lines = []
        lines.append(f"; Generated by AHK Script Builder - {self.app_name_edit.text()}")
        lines.append("; Keep the ACTION comments if you want to reload this script into the builder.")
        lines.append("CoordMode, Mouse, Screen")
        lines.append("SetTitleMatchMode, 2")
        lines.append("")
        # Create a main hotkey so running the ahk will execute the sequence when F9 pressed
        lines.append("; Press F9 to run the sequence")
        lines.append("F9::")
        for a in self.actions:
            lines.append(a.to_comment())
            # indent commands under hotkey
            for cmdline in a.to_ahk().split('\n'):
                if cmdline.strip():
                    lines.append('    ' + cmdline)
        lines.append('return')
        # Also include a simple run-at-start option: uncomment if desired
        content = '\n'.join(lines)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        QMessageBox.information(self, "Exported", f"AHK script exported to {path}\nIt will run the actions when you press F9 while the .ahk is running.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
