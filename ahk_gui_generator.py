"""
AHK Script Builder (PySide6 + pyautogui)

Features:
- GUI to create named actions (Click, Sleep, Run, KeyPress, WaitWindow)
- Capture mouse coordinates for Click actions
- Save/load project (JSON)
- Export .ahk file that contains both actionable AHK commands and structured comment lines
- Load an .ahk file if it was created by this tool (parses comment ACTION lines)
- Choose custom hotkey for exported script
- Reorder actions (Move Up / Move Down)
- Run action: pick installed Start Menu programs, browse EXE, add params, start mode
"""
import sys
import json
import os
import time
from enum import Enum
from dataclasses import dataclass, asdict
from typing import List, Optional

from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QMessageBox, QFileDialog,
    QDialog, QFormLayout, QComboBox, QSpinBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QToolBar
)
from PySide6.QtGui import QAction
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
    # Extended Run options
    mode: str = "Normal"       # Normal / Minimized / Maximized / Hidden
    params: str = ""          # additional command-line parameters

    def to_comment(self) -> str:
        # Structured comment line used by the GUI to reload
        if self.type == ActionType.CLICK:
            return f";ACTION:Click;Name={self.name};X={self.x};Y={self.y}"
        if self.type == ActionType.SLEEP:
            return f";ACTION:Sleep;Name={self.name};MS={self.ms}"
        if self.type == ActionType.RUN:
            # include params and mode in comment for round-tripping
            safe_params = self.params.replace(';', ',')
            return f";ACTION:Run;Name={self.name};Cmd={self.command};Params={safe_params};Mode={self.mode}"
        if self.type == ActionType.KEYPRESS:
            return f";ACTION:KeyPress;Name={self.name};Key={self.command}"
        if self.type == ActionType.WAITWINDOW:
            return f";ACTION:WaitWindow;Name={self.name};Title={self.command}"
        return f";ACTION:Unknown;Name={self.name}"

    def to_ahk_lines(self) -> List[str]:
        # Return list of AHK command lines (without indentation)
        lines: List[str] = []
        if self.type == ActionType.CLICK:
            lines.append(f"MouseMove, {self.x}, {self.y}")
            lines.append("Click")
            return lines
        if self.type == ActionType.SLEEP:
            lines.append(f"Sleep, {self.ms}")
            return lines
        if self.type == ActionType.RUN:
            # Prepare param text and mode keyword
            param_text = f" {self.params}" if self.params else ""
            # Mode mapping to AHK Run, ,, mode
            mode_map = {
                "Normal": "",
                "Minimized": "Min",
                "Maximized": "Max",
                "Hidden": "Hide"
            }
            ahk_mode = mode_map.get(self.mode, "")
            # If command contains spaces or is a path, wrap it in quotes
            cmd = self.command
            # If user selected a Start Menu entry (not full path) we write it as-is (AHK Run will find registered apps)
            # But if it's a path we ensure quoting:
            if os.path.exists(cmd) and (" " in cmd or '"' not in cmd):
                cmd_wrapped = f"\"{cmd}\""
            else:
                cmd_wrapped = cmd
            if ahk_mode:
                lines.append(f"Run, {cmd_wrapped}{param_text}, , {ahk_mode}")
            else:
                lines.append(f"Run, {cmd_wrapped}{param_text}")
            return lines
        if self.type == ActionType.KEYPRESS:
            lines.append(f"Send, {self.command}")
            return lines
        if self.type == ActionType.WAITWINDOW:
            lines.append(f"WinWait, {self.command}, , 10")
            lines.append(f"WinActivate, {self.command}")
            return lines
        return lines


class ActionDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None, action: Optional[Action] = None):
        super().__init__(parent)
        self.setWindowTitle("Add / Edit Action")
        self.resize(480, 260)
        self._parent_main = parent if isinstance(parent, MainWindow) else None

        self.form_layout = QFormLayout()

        self.type_cb = QComboBox()
        for t in ActionType:
            self.type_cb.addItem(t.value)
        self.form_layout.addRow("Type:", self.type_cb)

        self.name_edit = QLineEdit()
        self.form_layout.addRow("Name:", self.name_edit)

        # For Click: coords
        coords_row = QHBoxLayout()
        self.x_spin = QSpinBox()
        self.x_spin.setRange(0, 99999)
        self.y_spin = QSpinBox()
        self.y_spin.setRange(0, 99999)
        coords_row.addWidget(QLabel("X:"))
        coords_row.addWidget(self.x_spin)
        coords_row.addWidget(QLabel("Y:"))
        coords_row.addWidget(self.y_spin)
        self.capture_btn = QPushButton("Capture Mouse Position")
        self.capture_btn.clicked.connect(self.capture_position)
        coords_row.addWidget(self.capture_btn)
        self.form_layout.addRow("Coordinates:", coords_row)

        # Sleep ms
        self.ms_spin = QSpinBox()
        self.ms_spin.setRange(0, 6000000)
        self.ms_spin.setSingleStep(100)
        self.form_layout.addRow("Milliseconds:", self.ms_spin)

        # Command field (Run / KeyPress / WaitWindow)
        self.cmd_edit = QLineEdit()

        # Installed programs dropdown + browse
        self.program_cb = QComboBox()
        self.program_cb.addItem("Select installed program...")
        if self._parent_main:
            for name in self._parent_main.get_installed_programs():
                self.program_cb.addItem(name)

        self.browse_btn = QPushButton("Browse EXE")
        self.browse_btn.clicked.connect(self.browse_exe)

        run_layout = QHBoxLayout()
        run_layout.addWidget(self.program_cb)
        run_layout.addWidget(self.browse_btn)

        self.program_cb.currentTextChanged.connect(self.program_selected)

        self.form_layout.addRow("Command / Program:", self.cmd_edit)
        self.form_layout.addRow("Installed programs:", run_layout)

        # Extended Run options
        self.mode_cb = QComboBox()
        self.mode_cb.addItems(["Normal", "Minimized", "Maximized", "Hidden"])
        self.params_edit = QLineEdit()
        self.form_layout.addRow("Start mode:", self.mode_cb)
        self.form_layout.addRow("Extra parameters:", self.params_edit)

        # Buttons
        btn_layout = QHBoxLayout()
        self.ok_btn = QPushButton("OK")
        self.cancel_btn = QPushButton("Cancel")
        self.ok_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.ok_btn)
        btn_layout.addWidget(self.cancel_btn)
        self.form_layout.addRow(btn_layout)

        self.setLayout(self.form_layout)

        # Signals
        self.type_cb.currentTextChanged.connect(self.type_changed)

        if action:
            self.load_action(action)
        else:
            self.type_changed(self.type_cb.currentText())

    def program_selected(self, text: str):
        if text and text != "Select installed program...":
            # Fill command field with program name; user can override or browse
            self.cmd_edit.setText(text)

    def browse_exe(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose executable", "", "Executable (*.exe);;All files (*)")
        if path:
            self.cmd_edit.setText(path)

    def load_action(self, action: Action):
        self.type_cb.setCurrentText(action.type.value)
        self.name_edit.setText(action.name)
        self.x_spin.setValue(action.x)
        self.y_spin.setValue(action.y)
        self.ms_spin.setValue(action.ms)
        self.cmd_edit.setText(action.command)
        self.mode_cb.setCurrentText(action.mode)
        self.params_edit.setText(action.params)
        self.type_changed(action.type.value)

    def type_changed(self, text: str):
        t = ActionType(text)
        # coordinates visible/enabled for Click
        coords_visible = (t == ActionType.CLICK)
        self.x_spin.setEnabled(coords_visible)
        self.y_spin.setEnabled(coords_visible)
        self.capture_btn.setEnabled(coords_visible)

        # sleep enabled for Sleep
        self.ms_spin.setEnabled(t == ActionType.SLEEP)

        # command enabled for Run/KeyPress/WaitWindow
        self.cmd_edit.setEnabled(t in (ActionType.RUN, ActionType.KEYPRESS, ActionType.WAITWINDOW))

        # show program dropdown/browse and extended run only for RUN
        show_run = (t == ActionType.RUN)
        self.program_cb.setVisible(show_run)
        self.browse_btn.setVisible(show_run)
        self.mode_cb.setVisible(show_run)
        self.params_edit.setVisible(show_run)

    def capture_position(self):
        QMessageBox.information(self, "Capture position",
                                "After you click OK, move the mouse to the target position and press OK again to capture the coordinates.")
        res = QMessageBox.question(self, "Ready to capture", "Move mouse to target position and press OK to capture current mouse position.", QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
        if res == QMessageBox.StandardButton.Ok:
            time.sleep(0.2)
            x, y = pyautogui.position()
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
            command=self.cmd_edit.text(),
            mode=self.mode_cb.currentText(),
            params=self.params_edit.text()
        )


class MainWindow(QMainWindow):
    def get_installed_programs(self) -> List[str]:
        """Return a clean sorted list of installed application names from Start Menu."""
        programs = []
        start_menu_paths = [
            os.path.expandvars(r"%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs"),
            os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs"),
        ]
        for base in start_menu_paths:
            if os.path.exists(base):
                for root, dirs, files in os.walk(base):
                    for f in files:
                        if f.lower().endswith(".lnk"):
                            name = f[:-4]
                            programs.append(name)
        return sorted(set(programs))

    def __init__(self):
        super().__init__()
        self.setWindowTitle("AHK Script Builder")
        self.resize(900, 520)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        top_layout = QHBoxLayout()
        layout.addLayout(top_layout)

        top_layout.addWidget(QLabel("Application name (for comments / file):"))
        self.app_name_edit = QLineEdit("MyScript")
        top_layout.addWidget(self.app_name_edit)

        top_layout.addWidget(QLabel("Hotkey (AHK format, e.g. ^!z or #+F12):"))
        self.hotkey_edit = QLineEdit("^!z")
        top_layout.addWidget(self.hotkey_edit)

        top_layout.addStretch()

        # Add/Edit/Remove buttons
        btns_layout = QHBoxLayout()
        layout.addLayout(btns_layout)

        self.new_btn = QPushButton("New Action")
        self.new_btn.clicked.connect(self.add_action)
        btns_layout.addWidget(self.new_btn)

        self.edit_btn = QPushButton("Edit")
        self.edit_btn.clicked.connect(self.edit_action)
        btns_layout.addWidget(self.edit_btn)

        self.remove_btn = QPushButton("Remove")
        self.remove_btn.clicked.connect(self.remove_action)
        btns_layout.addWidget(self.remove_btn)

        # Actions table (primary UI)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["#", "Type", "Name", "Params / Info", "Preview Comment"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(self.table.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(self.table.SelectionMode.SingleSelection)
        layout.addWidget(self.table)

        # Reorder buttons
        reorder_layout = QHBoxLayout()
        self.move_up_btn = QPushButton("Move Up")
        self.move_up_btn.clicked.connect(self.move_action_up)
        reorder_layout.addWidget(self.move_up_btn)
        self.move_down_btn = QPushButton("Move Down")
        self.move_down_btn.clicked.connect(self.move_action_down)
        reorder_layout.addWidget(self.move_down_btn)
        reorder_layout.addStretch()
        layout.addLayout(reorder_layout)

        # Save/Load/Export buttons
        bottom_layout = QHBoxLayout()
        layout.addLayout(bottom_layout)

        self.load_project_btn = QPushButton("Load .ahk or project")
        self.load_project_btn.clicked.connect(self.load_file)
        bottom_layout.addWidget(self.load_project_btn)

        self.save_project_btn = QPushButton("Save project (.json)")
        self.save_project_btn.clicked.connect(self.save_project)
        bottom_layout.addWidget(self.save_project_btn)

        self.export_btn = QPushButton("Export .ahk")
        self.export_btn.clicked.connect(self.export_ahk)
        bottom_layout.addWidget(self.export_btn)

        bottom_layout.addStretch()

        # Toolbar
        toolbar = QToolBar()
        self.addToolBar(toolbar)
        toolbar.addAction(QAction("New", self, triggered=self.add_action))
        toolbar.addAction(QAction("Open", self, triggered=self.load_file))
        toolbar.addAction(QAction("Save", self, triggered=self.save_project))
        toolbar.addAction(QAction("Export .ahk", self, triggered=self.export_ahk))

        # Internal actions storage (Python list)
        self.actions: List[Action] = []

    # ----- action list management -----
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

    def move_action_up(self):
        row = self.table.currentRow()
        if row > 0:
            self.actions[row - 1], self.actions[row] = self.actions[row], self.actions[row - 1]
            self.refresh_table()
            self.table.selectRow(row - 1)

    def move_action_down(self):
        row = self.table.currentRow()
        if 0 <= row < len(self.actions) - 1:
            self.actions[row + 1], self.actions[row] = self.actions[row], self.actions[row + 1]
            self.refresh_table()
            self.table.selectRow(row + 1)

    def refresh_table(self):
        self.table.setRowCount(len(self.actions))
        for i, a in enumerate(self.actions):
            item_index = QTableWidgetItem(str(i + 1))
            item_type = QTableWidgetItem(a.type.value)
            item_name = QTableWidgetItem(a.name)
            if a.type == ActionType.CLICK:
                params = f"X={a.x}, Y={a.y}"
            elif a.type == ActionType.SLEEP:
                params = f"MS={a.ms}"
            elif a.type == ActionType.RUN:
                params = f"Cmd={a.command} Params={a.params} Mode={a.mode}"
            elif a.type == ActionType.KEYPRESS:
                params = f"Key={a.command}"
            elif a.type == ActionType.WAITWINDOW:
                params = f"Title={a.command}"
            else:
                params = ""
            item_params = QTableWidgetItem(params)
            item_preview = QTableWidgetItem(a.to_comment())
            self.table.setItem(i, 0, item_index)
            self.table.setItem(i, 1, item_type)
            self.table.setItem(i, 2, item_name)
            self.table.setItem(i, 3, item_params)
            self.table.setItem(i, 4, item_preview)
        # Resize to contents a little
        self.table.resizeRowsToContents()

    # ----- persistence -----
    def save_project(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save project", os.path.join(APP_DIR, "project.json"), "JSON Files (*.json)")
        if not path:
            return
        data = {
            'app_name': self.app_name_edit.text(),
            'hotkey': self.hotkey_edit.text(),
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
            self.hotkey_edit.setText(data.get('hotkey', '^!z'))
            self.actions = []
            for ad in data.get('actions', []):
                # ensure keys exist (backwards compatible)
                a = Action(
                    type=ActionType(ad['type']),
                    name=ad.get('name', ad.get('type', 'Action')),
                    x=ad.get('x', 0),
                    y=ad.get('y', 0),
                    ms=ad.get('ms', 0),
                    command=ad.get('command', ''),
                    mode=ad.get('mode', 'Normal'),
                    params=ad.get('params', '')
                )
                self.actions.append(a)
            self.refresh_table()
            QMessageBox.information(self, "Loaded", f"Project loaded from {path}")
            return

        # If .ahk, try to parse structured comment lines we wrote earlier
        if path.lower().endswith('.ahk'):
            actions: List[Action] = []
            app_name = os.path.splitext(os.path.basename(path))[0]
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    line = line.strip()
                    if line.startswith(';ACTION:'):
                        try:
                            rest = line[len(';ACTION:'):]
                            parts = rest.split(';')
                            typ = parts[0]
                            kv = {}
                            for p in parts[1:]:
                                if '=' in p:
                                    k, v = p.split('=', 1)
                                    kv[k] = v
                            if typ == 'Click':
                                a = Action(ActionType.CLICK, kv.get('Name', 'Click'), int(kv.get('X', 0)), int(kv.get('Y', 0)))
                            elif typ == 'Sleep':
                                a = Action(ActionType.SLEEP, kv.get('Name', 'Sleep'), ms=int(kv.get('MS', 0)))
                            elif typ == 'Run':
                                a = Action(ActionType.RUN, kv.get('Name', 'Run'), command=kv.get('Cmd', ''), params=kv.get('Params', ''), mode=kv.get('Mode', 'Normal'))
                            elif typ == 'KeyPress':
                                a = Action(ActionType.KEYPRESS, kv.get('Name', 'KeyPress'), command=kv.get('Key', ''))
                            elif typ == 'WaitWindow':
                                a = Action(ActionType.WAITWINDOW, kv.get('Name', 'WaitWindow'), command=kv.get('Title', ''))
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

    # ----- export to .ahk -----
    def export_ahk(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export .ahk", os.path.join(APP_DIR, f"{self.app_name_edit.text()}.ahk"), "AHK Scripts (*.ahk)")
        if not path:
            return
        hotkey = self.hotkey_edit.text().strip() or "F9"

        lines: List[str] = []
        lines.append(f"; Generated by AHK Script Builder - {self.app_name_edit.text()}")
        lines.append("; Keep the ACTION comments if you want to reload this script into the builder.")
        lines.append("CoordMode, Mouse, Screen")
        lines.append("SetTitleMatchMode, 2")
        lines.append("")
        lines.append(f"; Hotkey to run the sequence: {hotkey}")
        lines.append(f"{hotkey}::")
        for a in self.actions:
            # comment line
            lines.append(a.to_comment())
            # AHK commands indented
            for l in a.to_ahk_lines():
                lines.append("    " + l)
        lines.append("return")
        content = "\n".join(lines)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        QMessageBox.information(self, "Exported", f"AHK script exported to {path}\nIt will run the actions when you press {hotkey} while the .ahk is running.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
