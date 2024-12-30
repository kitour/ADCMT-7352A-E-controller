import sys
import time
import csv
import pyvisa
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QMessageBox, QMainWindow, QStackedWidget, QRadioButton, QButtonGroup,
    QGridLayout, QScrollArea, QLineEdit, QFileDialog, QGroupBox, QSizePolicy, QDialog
)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QFont
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import MultipleLocator
import matplotlib
import matplotlib.font_manager as font_manager
from multiprocessing import Event, Manager
import math
from threading import Thread

# 利用可能なフォントから日本語フォントを探す
available_fonts = [f.name for f in font_manager.fontManager.ttflist]
japanese_fonts = [name for name in available_fonts if
                  any(word in name for word in ['Yu Gothic', 'Meiryo', 'MS Gothic', 'Noto Sans CJK JP'])]

if japanese_fonts:
    matplotlib.rcParams['font.family'] = japanese_fonts
else:
    matplotlib.rcParams['font.family'] = ['sans-serif']

rm = pyvisa.ResourceManager()


def format_si_unit(value, unit='Ω'):
    if value == 'Overload' or value is None or math.isnan(value):
        return "Overload"

    si_prefixes = {
        -9: 'n',
        -6: 'μ',
        -3: 'm',
        0: '',
        3: 'k',
        6: 'M',
        9: 'G',
    }

    if value == 0:
        exponent = 0
    else:
        exponent = int(math.floor(math.log10(abs(value))))
        exponent = exponent - (exponent % 3)

    min_exponent = min(si_prefixes.keys())
    max_exponent = max(si_prefixes.keys())
    exponent = max(min(exponent, max_exponent), min_exponent)

    scaled_value = value / (10 ** exponent)
    prefix = si_prefixes[exponent]
    return f"{scaled_value:.6g} {prefix}{unit}"


class MeasurementClass(Thread):
    def __init__(self, command_queue, data_list, resource_name, stop_event, is_ready_event):
        super().__init__()
        self.dmm = None
        self.command_queue = command_queue
        self.data_list = data_list
        self.resource_name = resource_name
        self.stop_event = stop_event
        self.is_ready_event = is_ready_event

    def run(self):
        rm = pyvisa.ResourceManager()
        try:
            self.dmm = rm.open_resource(self.resource_name)
            self.dmm.timeout = 10000
            self.dmm.read_termination = '\n'
            self.dmm.write_termination = '\n'
        except Exception as e:
            print(f"MeasurementClass: 機器のオープンに失敗しました: {e}")
            return

        start_time = time.time()

        while not self.stop_event.is_set():
            self.check_commands()
            try:
                ach_value, bch_value = self.read_measurement()
                if ach_value is not None:
                    timestamp = time.time() - start_time
                    self.data_list.append((timestamp, ach_value, bch_value))
                    if not self.is_ready_event.is_set():
                        self.is_ready_event.set()
            except Exception as e:
                print(f"MeasurementClass: 例外が発生しました: {e}")
            time.sleep(0.005)

        self.dmm.close()
        print("MeasurementClass: プロセスが終了しました。")

    def check_commands(self):
        while not self.command_queue.empty():
            command = self.command_queue.get()
            if command == "STOP":
                self.stop_event.set()
            elif command.startswith("SEND "):
                cmd = command[5:]
                self.send_command(cmd)
            elif command == "TRIGGER":
                self.send_command("*TRG")

    def send_command(self, command):
        try:
            self.dmm.write("*CLS")
            self.dmm.clear()
            time.sleep(0.2)

            self.dmm.write(command)
            time.sleep(0.2)

            self.dmm.write("*OPC?")
            while True:
                try:
                    ready = self.dmm.read().strip()
                    if ready == "1":
                        break
                except pyvisa.errors.VisaIOError:
                    time.sleep(0.2)
        except Exception as e:
            print(f"MeasurementClass: コマンドの送信中にエラーが発生しました: {e}")

    def read_measurement(self):
        try:
            measurement = self.dmm.read().strip()
            parts = measurement.split(",")
            ach_value = None
            bch_value = None

            for part in parts:
                part = part.strip()
                if not part:
                    continue

                tokens = part.split()
                if len(tokens) < 2:
                    continue

                prefix = tokens[0]
                value_str = tokens[1]

                if prefix.endswith('_'):
                    try:
                        value = float(value_str)
                        if value == -0.0:
                            value = 0.0
                    except ValueError:
                        continue
                elif prefix.endswith('O'):
                    value = 'Overload'
                else:
                    continue

                if ach_value is None:
                    ach_value = value
                elif bch_value is None:
                    bch_value = value

            if ach_value is None:
                return None, None

            return ach_value, bch_value
        except pyvisa.errors.VisaIOError:
            return None, None


class DeviceSelectionPage(QWidget):
    def __init__(self, resources):
        super().__init__(parent=None)
        self.resources = resources
        self.selected_resource = None

        layout = QVBoxLayout()
        label = QLabel("以下の機器が見つかりました:")
        layout.addWidget(label)

        self.combo = QComboBox(parent=None)
        self.combo.addItems(resources)
        layout.addWidget(self.combo)

        self.next_button = QPushButton("次へ")
        self.next_button.setFixedSize(100, 40)
        layout.addWidget(self.next_button, alignment=Qt.AlignRight)

        self.setLayout(layout)


class DMMSetupPage(QWidget):
    def __init__(self, jig_modes):
        super().__init__(parent=None)
        self.jig_modes = jig_modes
        self.measurement_options = [
            ("F1", "直流電圧測定 (DCV-Ach)", "DCV", "V"),
            ("F2", "交流電圧測定 (ACV-Ach)", "ACV", "V"),
            ("F3", "抵抗測定 (2WΩ-Ach)", "OHM", "Ω"),
            ("F5", "直流電流測定 (DCI-Ach)", "DCI", "A"),
            ("F6", "交流電流測定 (ACI-Ach)", "ACI", "A"),
            ("F7", "交流電圧測定 (AC+DC結合)", "ACVDCV", "V"),
            ("F8", "交流電流測定 (AC+DC結合)", "ACIDCI", "A"),
            ("F12", "Bch 直流電圧測定 (DCV-Bch)", "BDV", "V"),
            ("F13", "ダイオード測定 (DIODE-Ach)", "DIODE", "V"),
            ("F20", "ローパワー2WΩ(LP-2W-Ach)", "LP2W", "Ω"),
            ("F22", "導通テスト (CONT-Ach)", "CONT", "Ω"),
            ("F35", "Bch 直流電流測定 (DCI-Bch)", "BDI", "A"),
            ("F36", "Bch 交流電流測定 (ACI-Bch)", "BAI", "A"),
            ("F37", "Bch 交流電流測定 (AC+DC結合)", "BAIDCI", "A"),
            ("F40", "温度測定 (TEMP)", "TEMP", "℃"),
            ("F50", "周波数測定 (FREQ-Ach)", "FREQ", "Hz"),
            ("DE0", "OFF", "", ""),
        ]
        self.range_dict = {
            "F1": [
                ("R0", "AUTO"),
                ("R3", "200 mV"),
                ("R4", "2 V"),
                ("R5", "20 V"),
                ("R6", "200 V"),
                ("R7", "1000 V"),
            ],
            "F2": [
                ("R0", "AUTO"),
                ("R3", "200 mV"),
                ("R4", "2 V"),
                ("R5", "20 V"),
                ("R6", "200 V"),
                ("R7", "700 V"),
            ],
            "F3": [
                ("R0", "AUTO"),
                ("R3", "200 Ω"),
                ("R4", "2 kΩ"),
                ("R5", "20 kΩ"),
                ("R6", "200 kΩ"),
                ("R7", "2 MΩ"),
                ("R8", "20 MΩ"),
                ("R9", "200 MΩ"),
            ],
            "F5": [
                ("R0", "AUTO"),
                ("R1", "2000 nA"),
                ("R2", "20 μA"),
                ("R3", "200 μA"),
                ("R4", "2 mA"),
                ("R5", "20 mA"),
                ("R6", "200 mA"),
                ("R7", "2000 mA"),
            ],
            "F6": [
                ("R0", "AUTO"),
                ("R3", "200 μA"),
                ("R4", "2 mA"),
                ("R5", "20 mA"),
                ("R6", "200 mA"),
                ("R7", "2000 mA"),
            ],
            "F7": [
                ("R0", "AUTO"),
                ("R3", "200 mV"),
                ("R4", "2 V"),
                ("R5", "20 V"),
                ("R6", "200 V"),
                ("R7", "700 V"),
            ],
            "F8": [
                ("R0", "AUTO"),
                ("R3", "200 μA"),
                ("R4", "2 mA"),
                ("R5", "20 mA"),
                ("R6", "200 mA"),
                ("R7", "2000 mA"),
            ],
            "F12": [
                ("R0", "AUTO"),
                ("R3", "200 mV"),
                ("R4", "2 V"),
                ("R5", "20 V"),
                ("R6", "200 V"),
            ],
            "F13": [
                ("R0", "AUTO"),
                ("R3", "200 Ω"),
                ("R4", "2 kΩ"),
                ("R5", "20 kΩ"),
                ("R6", "200 kΩ"),
                ("R7", "2 MΩ"),
                ("R8", "20 MΩ"),
                ("R9", "200 MΩ"),
            ],
            "F20": [
                ("R0", "AUTO"),
                ("R3", "200 Ω"),
                ("R4", "2 kΩ"),
                ("R5", "20 kΩ"),
                ("R6", "200 kΩ"),
                ("R7", "2 MΩ"),
                ("R8", "20 MΩ"),
            ],
            "F22": [
                ("R0", "AUTO"),
                ("R3", "200 Ω"),
                ("R4", "2 kΩ"),
                ("R5", "20 kΩ"),
                ("R6", "200 kΩ"),
                ("R7", "2 MΩ"),
                ("R8", "20 MΩ"),
                ("R9", "200 MΩ"),
            ],
            "F35": [
                ("R8", "10 A"),
            ],
            "F36": [
                ("R8", "10 A"),
            ],
            "F37": [
                ("R8", "10 A"),
            ]
        }

        main_layout = QVBoxLayout()
        font = QFont()
        font.setPointSize(14)
        self.setFont(font)

        mode_selection_group = QGroupBox("測定モードを選択してください")
        mode_selection_layout = QHBoxLayout()
        self.mode_group = QButtonGroup(parent=None)
        self.normal_mode_radio = QRadioButton("通常測定モード")
        self.jig_mode_radio = QRadioButton("ジグ使用計測モード")
        self.mode_group.addButton(self.normal_mode_radio, id=0)
        self.mode_group.addButton(self.jig_mode_radio, id=1)
        self.normal_mode_radio.setChecked(True)

        mode_selection_layout.addWidget(self.normal_mode_radio)
        mode_selection_layout.addWidget(self.jig_mode_radio)
        mode_selection_group.setLayout(mode_selection_layout)
        main_layout.addWidget(mode_selection_group)

        self.jig_selection_group = QGroupBox("ジグ測定モードを選択してください")
        jig_selection_layout = QHBoxLayout()
        self.jig_selection_combo = QComboBox(parent=None)
        for jig_mode in self.jig_modes:
            self.jig_selection_combo.addItem(jig_mode[0])
        jig_selection_layout.addWidget(self.jig_selection_combo)
        self.jig_selection_group.setLayout(jig_selection_layout)
        self.jig_selection_group.hide()
        main_layout.addWidget(self.jig_selection_group)

        scroll = QScrollArea(parent=None)
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget(parent=None)
        scroll.setWidget(scroll_widget)
        scroll_layout = QVBoxLayout(scroll_widget)

        dsp1_group = QGroupBox("DSP1 の設定")
        dsp1_layout = QVBoxLayout()
        dsp1_options_layout = QGridLayout()

        self.dsp1_option_group = QButtonGroup(parent=None)
        for idx, (key, desc, prefix, unit) in enumerate(self.measurement_options):
            if key == "DE0":
                continue
            btn = QRadioButton(desc)
            self.dsp1_option_group.addButton(btn, idx)
            dsp1_options_layout.addWidget(btn, idx // 4, idx % 4)
        if self.dsp1_option_group.buttons():
            self.dsp1_option_group.buttons()[0].setChecked(True)
        dsp1_layout.addLayout(dsp1_options_layout)

        dsp1_range_label = QLabel("レンジを選択してください:")
        dsp1_layout.addWidget(dsp1_range_label)
        self.dsp1_range_group = QButtonGroup(parent=None)
        self.dsp1_range_layout = QGridLayout()
        dsp1_layout.addLayout(self.dsp1_range_layout)
        self.update_dsp1_ranges()

        self.dsp1_option_group.buttonClicked.connect(self.update_dsp1_ranges)
        dsp1_group.setLayout(dsp1_layout)
        scroll_layout.addWidget(dsp1_group)

        dsp2_group = QGroupBox("DSP2 の設定")
        dsp2_layout = QVBoxLayout()
        dsp2_options_layout = QGridLayout()

        self.dsp2_option_group = QButtonGroup(parent=None)
        for idx, (key, desc, prefix, unit) in enumerate(self.measurement_options):
            btn = QRadioButton(desc)
            self.dsp2_option_group.addButton(btn, idx)
            dsp2_options_layout.addWidget(btn, idx // 4, idx % 4)
        if self.dsp2_option_group.buttons():
            self.dsp2_option_group.buttons()[0].setChecked(False)
        dsp2_layout.addLayout(dsp2_options_layout)

        dsp2_range_label = QLabel("レンジを選択してください:")
        dsp2_layout.addWidget(dsp2_range_label)
        self.dsp2_range_group = QButtonGroup(parent=None)
        self.dsp2_range_layout = QGridLayout()
        dsp2_layout.addLayout(self.dsp2_range_layout)
        self.update_dsp2_ranges()

        self.dsp2_option_group.buttonClicked.connect(self.update_dsp2_ranges)
        dsp2_group.setLayout(dsp2_layout)
        scroll_layout.addWidget(dsp2_group)

        trigger_sampling_layout = QVBoxLayout()

        trigger_group = QGroupBox("トリガーモード")
        trigger_layout = QHBoxLayout()
        self.trigger_group = QButtonGroup(parent=None)
        self.trigger_radio1 = QRadioButton("IMMEDIATE (TRS0)")
        self.trigger_radio1.setChecked(True)
        self.trigger_radio3 = QRadioButton("BUS (TRS3)")
        self.trigger_group.addButton(self.trigger_radio1)
        self.trigger_group.addButton(self.trigger_radio3)
        trigger_layout.addWidget(self.trigger_radio1)
        trigger_layout.addWidget(self.trigger_radio3)
        trigger_group.setLayout(trigger_layout)

        sampling_group = QGroupBox("サンプリングレート")
        sampling_layout = QHBoxLayout()
        self.sampling_group = QButtonGroup(parent=None)
        self.sampling_radio1 = QRadioButton("FAST (PR1)")
        self.sampling_radio1.setChecked(True)
        self.sampling_radio2 = QRadioButton("MED (PR2)")
        self.sampling_radio3 = QRadioButton("SLOW1 (PR3)")
        self.sampling_radio4 = QRadioButton("SLOW2 (PR4)")
        self.sampling_group.addButton(self.sampling_radio1)
        self.sampling_group.addButton(self.sampling_radio2)
        self.sampling_group.addButton(self.sampling_radio3)
        self.sampling_group.addButton(self.sampling_radio4)
        sampling_layout.addWidget(self.sampling_radio1)
        sampling_layout.addWidget(self.sampling_radio2)
        sampling_layout.addWidget(self.sampling_radio3)
        sampling_layout.addWidget(self.sampling_radio4)
        sampling_group.setLayout(sampling_layout)

        auto_zero_group = QGroupBox("オートゼロ設定")
        auto_zero_layout = QHBoxLayout()
        self.auto_zero_group = QButtonGroup(parent=None)
        self.auto_zero_radio_off = QRadioButton("OFF")
        self.auto_zero_radio_on = QRadioButton("ON")
        self.auto_zero_radio_once = QRadioButton("ONCE (一度実行後、OFF)")
        self.auto_zero_group.addButton(self.auto_zero_radio_off, id=0)
        self.auto_zero_group.addButton(self.auto_zero_radio_on, id=1)
        self.auto_zero_group.addButton(self.auto_zero_radio_once, id=2)
        self.auto_zero_radio_off.setChecked(True)
        auto_zero_layout.addWidget(self.auto_zero_radio_off)
        auto_zero_layout.addWidget(self.auto_zero_radio_on)
        auto_zero_layout.addWidget(self.auto_zero_radio_once)
        auto_zero_group.setLayout(auto_zero_layout)

        trigger_auto_zero_layout = QHBoxLayout()
        trigger_auto_zero_layout.addWidget(trigger_group)
        trigger_auto_zero_layout.addWidget(auto_zero_group)

        trigger_sampling_layout.addLayout(trigger_auto_zero_layout)
        trigger_sampling_layout.addWidget(sampling_group)
        scroll_layout.addLayout(trigger_sampling_layout)

        main_layout.addWidget(scroll)

        button_layout = QHBoxLayout()
        self.next_button = QPushButton("次へ")
        self.next_button.setFixedSize(100, 40)
        button_layout.addStretch()
        button_layout.addWidget(self.next_button)
        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)
        self.mode_group.buttonClicked.connect(self.handle_mode_selection)
        self.jig_selection_combo.currentIndexChanged.connect(self.handle_jig_mode_selection)

    def handle_mode_selection(self, button):
        mode_id = self.mode_group.id(button)
        if mode_id == 0:
            self.set_normal_mode()
        else:
            self.set_jig_mode()

    def handle_jig_mode_selection(self):
        self.update_dsp1_ranges()
        self.update_dsp2_ranges()

    def set_normal_mode(self):
        self.jig_selection_group.hide()
        for btn in self.dsp1_option_group.buttons():
            btn.setEnabled(True)
        for btn in self.dsp2_option_group.buttons():
            btn.setEnabled(True)

        if not any(btn.isChecked() for btn in self.dsp2_option_group.buttons()):
            for btn in self.dsp2_option_group.buttons():
                if btn.text() == "OFF":
                    btn.setChecked(True)
                    break

    def set_jig_mode(self):
        self.jig_selection_group.show()
        self.jig_selection_combo.setCurrentIndex(0)
        self.update_dsp1_ranges()
        self.update_dsp2_ranges()
        for btn in self.dsp1_option_group.buttons():
            btn.setEnabled(False)
        for btn in self.dsp2_option_group.buttons():
            btn.setEnabled(False)

    def update_dsp1_ranges(self):
        for btn in self.dsp1_range_group.buttons():
            self.dsp1_range_group.removeButton(btn)
            btn.deleteLater()
        for i in reversed(range(self.dsp1_range_layout.count())):
            widget = self.dsp1_range_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        mode_id = self.mode_group.checkedId()
        if mode_id == 1:
            jig_mode_name = self.jig_selection_combo.currentText()
            dsp1_key = None
            for jig_mode in self.jig_modes:
                if jig_mode[0] == jig_mode_name:
                    dsp1_key = jig_mode[1]
                    break
            if dsp1_key is None:
                QMessageBox.critical(self, "エラー", f"ジグ測定モード '{jig_mode_name}' のDSP1キーが見つかりません。")
                return
        else:
            idx = self.dsp1_option_group.checkedId()
            if idx == -1:
                QMessageBox.critical(self, "エラー", "DSP1の測定項目が選択されていません。")
                return
            dsp1_key = self.measurement_options[idx][0]

        if dsp1_key and dsp1_key in self.range_dict:
            ranges = self.range_dict[dsp1_key]
            for r_idx, (range_key, range_desc) in enumerate(ranges):
                btn = QRadioButton(range_desc)
                self.dsp1_range_group.addButton(btn, r_idx)
                self.dsp1_range_layout.addWidget(btn, r_idx // 8, r_idx % 8)
            if self.dsp1_range_group.buttons():
                self.dsp1_range_group.buttons()[0].setChecked(True)
        else:
            label = QLabel("レンジ設定はありません")
            self.dsp1_range_layout.addWidget(label)

    def update_dsp2_ranges(self):
        for btn in self.dsp2_range_group.buttons():
            self.dsp2_range_group.removeButton(btn)
            btn.deleteLater()
        for i in reversed(range(self.dsp2_range_layout.count())):
            widget = self.dsp2_range_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        mode_id = self.mode_group.checkedId()
        if mode_id == 1:
            jig_mode_name = self.jig_selection_combo.currentText()
            dsp2_key = None
            for jig_mode in self.jig_modes:
                if jig_mode[0] == jig_mode_name:
                    dsp2_key = jig_mode[2]
                    break
            if dsp2_key is None:
                QMessageBox.critical(self, "エラー", f"ジグ測定モード '{jig_mode_name}' のDSP2キーが見つかりません。")
                return
        else:
            idx = self.dsp2_option_group.checkedId()
            if idx == -1:
                label = QLabel("DSP2は選択されていません。")
                self.dsp2_range_layout.addWidget(label)
                return
            dsp2_key = self.measurement_options[idx][0]

        if dsp2_key and dsp2_key in self.range_dict:
            ranges = self.range_dict.get(dsp2_key, [])
            for r_idx, (range_key, range_desc) in enumerate(ranges):
                btn = QRadioButton(range_desc)
                self.dsp2_range_group.addButton(btn, r_idx)
                self.dsp2_range_layout.addWidget(btn, r_idx // 8, r_idx % 8)
            if self.dsp2_range_group.buttons():
                self.dsp2_range_group.buttons()[0].setChecked(True)
        else:
            label = QLabel("レンジ設定はありません")
            self.dsp2_range_layout.addWidget(label)

    def get_setup_commands(self):
        commands = []
        mode_id = self.mode_group.checkedId()
        if mode_id == 0:
            idx = self.dsp1_option_group.checkedId()
            if idx == -1:
                QMessageBox.critical(self, "エラー", "DSP1の測定項目が選択されていません。")
                return []
            dsp1_key, _, _, _ = self.measurement_options[idx]
            commands.append(f"DSP1,{dsp1_key}")

            if self.dsp1_range_group.buttons():
                range_idx = self.dsp1_range_group.checkedId()
                ranges = self.range_dict.get(dsp1_key, [])
                if 0 <= range_idx < len(ranges):
                    range_key, _ = ranges[range_idx]
                    commands.append(f"DSP1,{range_key}")

            idx = self.dsp2_option_group.checkedId()
            if idx == -1:
                commands.append("DE0")
            else:
                dsp2_key, _, _, _ = self.measurement_options[idx]
                if dsp2_key == "DE0":
                    commands.append("DE0")
                else:
                    commands.append(f"DSP2,{dsp2_key}")
                    if self.dsp2_range_group.buttons():
                        range_idx = self.dsp2_range_group.checkedId()
                        ranges = self.range_dict.get(dsp2_key, [])
                        if 0 <= range_idx < len(ranges):
                            range_key, _ = ranges[range_idx]
                            commands.append(f"DSP2,{range_key}")
        else:
            jig_mode_name = self.jig_selection_combo.currentText()
            dsp1_key = None
            dsp2_key = None
            for jig_mode in self.jig_modes:
                if jig_mode[0] == jig_mode_name:
                    dsp1_key = jig_mode[1]
                    dsp2_key = jig_mode[2]
                    break

            if dsp1_key and dsp1_key in self.range_dict:
                commands.append(f"DSP1,{dsp1_key}")
                ranges_dsp1 = self.range_dict.get(dsp1_key, [])
                range_idx = self.dsp1_range_group.checkedId()
                if 0 <= range_idx < len(ranges_dsp1):
                    range_key, _ = ranges_dsp1[range_idx]
                    commands.append(f"DSP1,{range_key}")
            else:
                QMessageBox.critical(self, "エラー", f"ジグ測定モード '{jig_mode_name}' のDSP1設定に誤りがあります。")
                return []

            if dsp2_key and dsp2_key in self.range_dict:
                commands.append(f"DSP2,{dsp2_key}")
                ranges_dsp2 = self.range_dict.get(dsp2_key, [])
                range_idx = self.dsp2_range_group.checkedId()
                if 0 <= range_idx < len(ranges_dsp2):
                    range_key, _ = ranges_dsp2[range_idx]
                    commands.append(f"DSP2,{range_key}")
            else:
                QMessageBox.critical(self, "エラー", f"ジグ測定モード '{jig_mode_name}' のDSP2設定に誤りがあります。")
                return []

        auto_zero_id = self.auto_zero_group.checkedId()
        if auto_zero_id == 0:
            commands.append("AZ0")
        elif auto_zero_id == 1:
            commands.append("AZ1")
        elif auto_zero_id == 2:
            commands.append("AZ2")

        if self.trigger_radio1.isChecked():
            commands.append("TRS0")
        elif self.trigger_radio3.isChecked():
            commands.append("TRS3")

        if self.sampling_radio1.isChecked():
            commands.append("PR1")
        elif self.sampling_radio2.isChecked():
            commands.append("PR2")
        elif self.sampling_radio3.isChecked():
            commands.append("PR3")
        elif self.sampling_radio4.isChecked():
            commands.append("PR4")

        return commands

    def get_trigger_mode(self):
        if self.trigger_radio1.isChecked():
            return "TRS0"
        elif self.trigger_radio3.isChecked():
            return "TRS3"

    def get_measurement_modes(self):
        mode_id = self.mode_group.checkedId()
        if mode_id == 0:
            idx = self.dsp1_option_group.checkedId()
            if idx == -1:
                dsp1_desc = ""
                dsp1_prefix = ""
                dsp1_unit = ""
            else:
                dsp1_desc = self.measurement_options[idx][1]
                dsp1_prefix = self.measurement_options[idx][2]
                dsp1_unit = self.measurement_options[idx][3]

            idx = self.dsp2_option_group.checkedId()
            if idx == -1:
                dsp2_desc = None
                dsp2_prefix = None
                dsp2_unit = None
            else:
                dsp2_key, dsp2_desc, dsp2_prefix, dsp2_unit = self.measurement_options[idx]
                if dsp2_key == "DE0":
                    dsp2_desc = None
                    dsp2_prefix = None
                    dsp2_unit = None

            return (dsp1_desc, dsp1_prefix, dsp1_unit), (dsp2_desc, dsp2_prefix, dsp2_unit)

        else:
            jig_mode_name = self.jig_selection_combo.currentText()
            dsp1_desc = ""
            dsp1_prefix = ""
            dsp1_unit = ""
            dsp2_desc = ""
            dsp2_prefix = ""
            dsp2_unit = ""
            for jig_mode in self.jig_modes:
                if jig_mode[0] == jig_mode_name:
                    dsp1_key = jig_mode[1]
                    dsp2_key = jig_mode[2]
                    dsp1_idx = self.get_measurement_option_index(dsp1_key)
                    dsp2_idx = self.get_measurement_option_index(dsp2_key)
                    if dsp1_idx != -1:
                        dsp1_desc = self.measurement_options[dsp1_idx][1]
                        dsp1_prefix = self.measurement_options[dsp1_idx][2]
                        dsp1_unit = self.measurement_options[dsp1_idx][3]
                    else:
                        QMessageBox.critical(self, "エラー",
                                             f"ジグ測定モード '{jig_mode_name}' のDSP1キーが見つかりません。")
                    if dsp2_idx != -1:
                        dsp2_desc = self.measurement_options[dsp2_idx][1]
                        dsp2_prefix = self.measurement_options[dsp2_idx][2]
                        dsp2_unit = self.measurement_options[dsp2_idx][3]
                    else:
                        QMessageBox.critical(self, "エラー",
                                             f"ジグ測定モード '{jig_mode_name}' のDSP2キーが見つかりません。")
                    break
            return (dsp1_desc, dsp1_prefix, dsp1_unit), (dsp2_desc, dsp2_prefix, dsp2_unit)

    def get_measurement_option_index(self, key):
        for idx, option in enumerate(self.measurement_options):
            if option[0] == key:
                return idx
        return -1


class ModeSelectionPage(QWidget):
    def __init__(self):
        super().__init__(parent=None)
        layout = QVBoxLayout()
        label = QLabel("モードを選択してください:")
        layout.addWidget(label)

        button_layout = QHBoxLayout()
        self.value_display_button = QPushButton("値表示モード")
        self.value_display_button.setFixedSize(200, 60)
        self.graph_display_button = QPushButton("グラフ表示モード")
        self.graph_display_button.setFixedSize(200, 60)
        button_layout.addWidget(self.value_display_button)
        button_layout.addWidget(self.graph_display_button)
        layout.addLayout(button_layout)

        self.reset_button = QPushButton("リセット")
        self.reset_button.setFixedSize(100, 40)
        layout.addWidget(self.reset_button, alignment=Qt.AlignRight)
        self.setLayout(layout)


class ValueDisplayPage(QWidget):
    def __init__(self):
        super().__init__(parent=None)
        self.calculated_unit = None
        layout = QVBoxLayout()
        self.value_label_ach = QLabel("---")
        self.value_label_bch = QLabel("---")
        self.value_label_calculated = QLabel("---")

        font = QFont()
        font.setPointSize(72)
        self.value_label_ach.setFont(font)
        self.value_label_bch.setFont(font)
        self.value_label_calculated.setFont(font)

        self.value_label_ach.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.value_label_bch.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.value_label_calculated.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        layout.addWidget(self.value_label_ach)
        layout.addWidget(self.value_label_bch)
        layout.addWidget(self.value_label_calculated)

        button_layout = QHBoxLayout()
        self.trigger_button = QPushButton("トリガー")
        self.switch_display_button = QPushButton("表示切替")
        self.reset_button = QPushButton("リセット")
        button_layout.addWidget(self.trigger_button)
        button_layout.addWidget(self.switch_display_button)
        button_layout.addWidget(self.reset_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.ach_desc = ""
        self.bch_desc = ""
        self.calculated_value_desc = ""
        self.jig_mode = False

    def set_measurement_mode_descriptions(self, ach_desc, bch_desc, jig_mode=False,
                                          calculated_value_desc="", calculated_unit=""):
        self.ach_desc = ach_desc
        self.bch_desc = bch_desc
        self.jig_mode = jig_mode
        self.calculated_value_desc = calculated_value_desc
        self.calculated_unit = calculated_unit

        if jig_mode:
            self.value_label_ach.hide()
            self.value_label_bch.hide()
            self.value_label_calculated.setText(f"{self.calculated_value_desc}: --- {self.calculated_unit}")
            self.value_label_calculated.show()
        else:
            self.value_label_calculated.hide()
            self.value_label_ach.setText(f"{self.ach_desc}: ---")
            if self.bch_desc:
                self.value_label_bch.setText(f"{self.bch_desc}: ---")
                self.value_label_bch.show()
            else:
                self.value_label_bch.hide()

    def update_values(self, ach_value, bch_value, calculated_value=None,
                      ach_unit="", bch_unit="", calculated_unit="", jig_mode=False):
        self.jig_mode = jig_mode
        if jig_mode:
            # 修正箇所: float('inf') をチェックに追加
            if calculated_value == 'Overload' or calculated_value is None or math.isnan(calculated_value) or math.isinf(calculated_value):
                calculated_text = f"{self.calculated_value_desc}: Overload"
            else:
                if calculated_unit:
                    formatted_value = format_si_unit(calculated_value, calculated_unit)
                else:
                    formatted_value = f"{calculated_value:.3f}"
                calculated_text = f"{self.calculated_value_desc}: {formatted_value}"
            self.value_label_calculated.setText(calculated_text)
        else:
            if ach_value == 'Overload' or math.isnan(ach_value):
                ach_text = f"{self.ach_desc}: Overload"
            else:
                formatted_value = format_si_unit(ach_value, ach_unit)
                ach_text = f"{self.ach_desc}: {formatted_value}"
            self.value_label_ach.setText(ach_text)

            if self.bch_desc:
                if bch_value == 'Overload' or math.isnan(bch_value):
                    bch_text = f"{self.bch_desc}: Overload"
                else:
                    formatted_value = format_si_unit(bch_value, bch_unit)
                    bch_text = f"{self.bch_desc}: {formatted_value}"
                self.value_label_bch.setText(bch_text)
                self.value_label_bch.show()
            else:
                self.value_label_bch.hide()

            self.value_label_calculated.hide()


class GraphDisplayPage(QWidget):
    def __init__(self):
        super().__init__(parent=None)
        self.bch_desc = None
        self.ach_desc = None
        self.line_bch = None
        self.line_ach = None
        self.line_calculated = None
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.ax_ach = None
        self.ax_bch = None
        self.ax_calculated = None

        self.value_label_ach = QLabel("---")
        self.value_label_bch = QLabel("---")
        self.value_label_calculated = QLabel("---")

        font = QFont()
        font.setPointSize(24)
        self.value_label_ach.setFont(font)
        self.value_label_bch.setFont(font)
        self.value_label_calculated.setFont(font)

        self.value_label_ach.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.value_label_bch.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        self.value_label_calculated.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        labels_layout = QHBoxLayout()
        labels_layout.setContentsMargins(0, 0, 0, 0)
        labels_layout.setSpacing(10)
        self.value_label_ach.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.value_label_bch.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.value_label_calculated.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        labels_layout.addWidget(self.value_label_ach, alignment=Qt.AlignLeft)
        labels_layout.addStretch(1)
        labels_layout.addWidget(self.value_label_bch, alignment=Qt.AlignCenter)
        labels_layout.addStretch(1)
        labels_layout.addWidget(self.value_label_calculated, alignment=Qt.AlignRight)
        main_layout.addLayout(labels_layout)

        self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        main_layout.addWidget(self.canvas, stretch=1)

        time_setting_layout = QHBoxLayout()
        time_setting_layout.setContentsMargins(0, 0, 0, 0)
        time_setting_layout.setSpacing(5)
        time_label = QLabel("時間軸の最大表示量（秒）:")
        self.time_input = QLineEdit("10")
        self.time_input.setFixedWidth(80)
        time_setting_layout.addWidget(time_label)
        time_setting_layout.addWidget(self.time_input)
        time_setting_layout.addStretch()
        main_layout.addLayout(time_setting_layout)

        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(5)

        self.trigger_button = QPushButton("トリガー")
        self.switch_display_button = QPushButton("表示切替")
        self.recording_button = QPushButton("記録開始")
        self.reset_graph_button = QPushButton("グラフリセット")
        self.reset_button = QPushButton("リセット")

        button_layout.addWidget(self.trigger_button)
        button_layout.addWidget(self.switch_display_button)
        button_layout.addWidget(self.recording_button)
        button_layout.addWidget(self.reset_graph_button)
        button_layout.addWidget(self.reset_button)
        button_layout.addStretch()

        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

        self.ach_data = []
        self.bch_data = []
        self.calculated_data = []
        self.time_data = []
        self.start_time = time.time()
        self.ach_unit = ""
        self.bch_unit = ""
        self.calculated_unit = ""
        self.calculated_value_desc = ""
        self.jig_mode = False

        self.recording = False
        self.recorded_data = []

        self.recording_button.clicked.connect(self.toggle_recording)
        self.reset_graph_button.clicked.connect(self.reset_graph)

    def load_data_from_list(self, data_list, jig_mode, jig_mode_name, measurement_mode_bch):
        ach_values = []
        bch_values = []
        calculated_values = []
        time_values = []

        for timestamp, ach_value, bch_value in data_list:
            time_values.append(timestamp)
            if jig_mode:
                if jig_mode_name in ("四端子抵抗測定A_V", "四端子抵抗測定B_V"):
                    if ach_value != 0 and not math.isnan(ach_value) and not math.isnan(bch_value):
                        calculated_value = abs(ach_value / bch_value)
                    else:
                        calculated_value = float('inf')  # 修正: 割り切れない場合はinf
                elif jig_mode_name == "hFE測定":
                    if ach_value != 0 and not math.isnan(ach_value) and not math.isnan(bch_value):
                        calculated_value = bch_value / ach_value
                    else:
                        calculated_value = float('inf')  # 修正
                elif jig_mode_name == "電力計測(小電流)":
                    if not math.isnan(ach_value) and not math.isnan(bch_value):
                        calculated_value = ach_value * bch_value
                    else:
                        calculated_value = float('inf')  # 修正
                elif jig_mode_name == "電力計測(大電流)":
                    if not math.isnan(ach_value) and not math.isnan(bch_value):
                        calculated_value = ach_value * bch_value
                    else:
                        calculated_value = float('inf')  # 修正
                else:
                    calculated_value = float('inf')  # 修正
                calculated_values.append(calculated_value)
            else:
                ach_values.append(ach_value)
                if measurement_mode_bch:
                    bch_values.append(bch_value)

        if jig_mode and calculated_values:
            self.update_graph(
                ach_values=[],
                bch_values=[],
                time_values=time_values,
                calculated_values=calculated_values
            )
        elif ach_values:
            self.update_graph(
                ach_values=ach_values,
                bch_values=bch_values if bch_values else None,
                time_values=time_values
            )

    def set_measurement_mode_descriptions(self, ach_desc, bch_desc, jig_mode=False,
                                          calculated_value_desc="", calculated_unit=""):
        self.ach_desc = ach_desc
        self.bch_desc = bch_desc
        self.jig_mode = jig_mode
        self.calculated_value_desc = calculated_value_desc
        self.calculated_unit = calculated_unit

        self.figure.clf()
        try:
            max_display_time = float(self.time_input.text())
        except ValueError:
            max_display_time = 10

        if self.jig_mode:
            self.ax_calculated = self.figure.add_subplot(1, 1, 1)
            self.ax_calculated.set_title(self.calculated_value_desc)
            self.line_calculated, = self.ax_calculated.plot([], [], color='red', linewidth=2)
            self.ax_calculated.grid(True, linestyle='--', color='gray', linewidth=0.5)
            self.ax_calculated.set_xlim(0, max_display_time)
            self.ax_calculated.set_xlabel("時間 (秒)")
            self.ax_calculated.set_ylabel(f"{self.calculated_value_desc} ({self.calculated_unit})")
            self.ax_calculated.ticklabel_format(useOffset=False, style='plain')
        else:
            if self.bch_desc:
                self.ax_ach = self.figure.add_subplot(2, 1, 1)
                self.ax_bch = self.figure.add_subplot(2, 1, 2)
                self.ax_ach.set_title(self.ach_desc)
                self.ax_bch.set_title(self.bch_desc)
                self.line_ach, = self.ax_ach.plot([], [], color='green', linewidth=2)
                self.line_bch, = self.ax_bch.plot([], [], color='blue', linewidth=2)
                self.ax_ach.grid(True, linestyle='--', color='gray', linewidth=0.5)
                self.ax_bch.grid(True, linestyle='--', color='gray', linewidth=0.5)
                self.ax_ach.set_xlim(0, max_display_time)
                self.ax_bch.set_xlim(0, max_display_time)
                self.ax_ach.set_xlabel("時間 (秒)")
                self.ax_ach.set_ylabel(f"{self.ach_desc} ({self.ach_unit})")
                self.ax_bch.set_xlabel("時間 (秒)")
                self.ax_bch.set_ylabel(f"{self.bch_desc} ({self.bch_unit})")
                self.ax_ach.ticklabel_format(useOffset=False, style='plain')
                self.ax_bch.ticklabel_format(useOffset=False, style='plain')
            else:
                self.ax_ach = self.figure.add_subplot(1, 1, 1)
                self.ax_ach.set_title(self.ach_desc)
                self.line_ach, = self.ax_ach.plot([], [], color='green', linewidth=2)
                self.ax_ach.grid(True, linestyle='--', color='gray', linewidth=0.5)
                self.ax_ach.set_xlim(0, max_display_time)
                self.ax_ach.set_xlabel("時間 (秒)")
                self.ax_ach.set_ylabel(f"{self.ach_desc} ({self.ach_unit})")
                self.ax_ach.ticklabel_format(useOffset=False, style='plain')

        self.figure.tight_layout()

    def update_graph(self, ach_values, bch_values, time_values, calculated_values=None):
        # 先にtime_dataにextend
        self.time_data.extend(time_values)

        # 空なら描画せずreturn
        if not self.time_data:
            self.canvas.draw()
            return

        if self.jig_mode and calculated_values:
            # 修正: float('inf') をフィルタリング
            filtered_time = []
            filtered_calculated = []
            for t, c in zip(time_values, calculated_values):
                if not math.isinf(c):
                    filtered_time.append(t)
                    filtered_calculated.append(c)
            self.calculated_data.extend(filtered_calculated)
            time_values = filtered_time
        else:
            self.ach_data.extend(ach_values)
            if self.bch_desc and bch_values:
                self.bch_data.extend(bch_values)

        try:
            max_display_time = float(self.time_input.text())
        except ValueError:
            max_display_time = 10

        if max_display_time > 0:
            min_time = max(0, self.time_data[-1] - max_display_time)
            max_time = self.time_data[-1]
        else:
            min_time = min(self.time_data)
            max_time = max(self.time_data)

        # フィルタリングしたデータのみを保持
        if self.jig_mode and calculated_values:
            indices = [i for i, t in enumerate(self.time_data) if min_time <= t <= max_time]
            self.time_data = [self.time_data[i] for i in indices]
            self.calculated_data = [self.calculated_data[i] for i in indices]
            time_data = [self.time_data[i] for i in indices]
            calculated_data = [self.calculated_data[i] for i in indices]
        else:
            indices = [i for i, t in enumerate(self.time_data) if min_time <= t <= max_time]
            self.time_data = [self.time_data[i] for i in indices]
            self.ach_data = [self.ach_data[i] for i in indices]
            if self.bch_desc and bch_values:
                self.bch_data = [self.bch_data[i] for i in indices]
            time_data = [self.time_data[i] for i in indices]
            ach_data = [self.ach_data[i] for i in indices]
            if self.bch_desc and bch_values:
                bch_data = [self.bch_data[i] for i in indices]

        if self.jig_mode and calculated_values:
            self.line_calculated.set_data(time_data, calculated_data)
            self.ax_calculated.set_xlim(min_time, max_time)
            self.ax_calculated.relim()
            self.ax_calculated.autoscale_view()

            y_min, y_max = self.ax_calculated.get_ylim()
            y_range = y_max - y_min if (y_max - y_min) != 0 else 1
            y_major_interval = y_range / 10
            self.ax_calculated.yaxis.set_major_locator(MultipleLocator(y_major_interval))
            self.ax_calculated.yaxis.set_minor_locator(MultipleLocator(y_major_interval / 5))

            x_range = max_time - min_time if (max_time - min_time) != 0 else 1
            x_major_interval = x_range / 10
            self.ax_calculated.xaxis.set_major_locator(MultipleLocator(x_major_interval))
            self.ax_calculated.xaxis.set_minor_locator(MultipleLocator(x_major_interval / 5))
            self.ax_calculated.grid(True, which='both', linestyle='--', color='gray', linewidth=0.5)
        else:
            if self.bch_desc:
                self.line_ach.set_data(time_data, ach_data)
                self.line_bch.set_data(time_data, bch_data)
                self.ax_ach.set_xlim(min_time, max_time)
                self.ax_bch.set_xlim(min_time, max_time)
                self.ax_ach.relim()
                self.ax_ach.autoscale_view()
                self.ax_bch.relim()
                self.ax_bch.autoscale_view()

                y_min, y_max = self.ax_ach.get_ylim()
                y_range = y_max - y_min if (y_max - y_min) != 0 else 1
                y_major_interval = y_range / 10
                self.ax_ach.yaxis.set_major_locator(MultipleLocator(y_major_interval))
                self.ax_ach.yaxis.set_minor_locator(MultipleLocator(y_major_interval / 5))

                y_min, y_max = self.ax_bch.get_ylim()
                y_range = y_max - y_min if (y_max - y_min) != 0 else 1
                y_major_interval = y_range / 10
                self.ax_bch.yaxis.set_major_locator(MultipleLocator(y_major_interval))
                self.ax_bch.yaxis.set_minor_locator(MultipleLocator(y_major_interval / 5))

                x_range = max_time - min_time if (max_time - min_time) != 0 else 1
                x_major_interval = x_range / 10
                self.ax_ach.xaxis.set_major_locator(MultipleLocator(x_major_interval))
                self.ax_ach.xaxis.set_minor_locator(MultipleLocator(x_major_interval / 5))
                self.ax_bch.xaxis.set_major_locator(MultipleLocator(x_major_interval))
                self.ax_bch.xaxis.set_minor_locator(MultipleLocator(x_major_interval / 5))
                self.ax_ach.grid(True, which='both', linestyle='--', color='gray', linewidth=0.5)
                self.ax_bch.grid(True, which='both', linestyle='--', color='gray', linewidth=0.5)
            else:
                self.line_ach.set_data(time_data, ach_data)
                self.ax_ach.set_xlim(min_time, max_time)
                self.ax_ach.relim()
                self.ax_ach.autoscale_view()

                y_min, y_max = self.ax_ach.get_ylim()
                y_range = y_max - y_min if (y_max - y_min) != 0 else 1
                y_major_interval = y_range / 10
                self.ax_ach.yaxis.set_major_locator(MultipleLocator(y_major_interval))
                self.ax_ach.yaxis.set_minor_locator(MultipleLocator(y_major_interval / 5))

                x_range = max_time - min_time if (max_time - min_time) != 0 else 1
                x_major_interval = x_range / 10
                self.ax_ach.xaxis.set_major_locator(MultipleLocator(x_major_interval))
                self.ax_ach.xaxis.set_minor_locator(MultipleLocator(x_major_interval / 5))
                self.ax_ach.grid(True, which='both', linestyle='--', color='gray', linewidth=0.5)

        self.canvas.draw()

        # ラベル表示
        if self.jig_mode and calculated_values:
            latest_value = calculated_values[-1] if calculated_values else float('inf')
            if latest_value == 'Overload' or latest_value is None or math.isnan(latest_value) or math.isinf(latest_value):
                calculated_text = f"{self.calculated_value_desc}: Overload"
            else:
                if self.calculated_unit:
                    formatted_value = format_si_unit(latest_value, self.calculated_unit)
                else:
                    formatted_value = f"{latest_value:.3f}"
                calculated_text = f"{self.calculated_value_desc}: {formatted_value}"
            self.value_label_calculated.setText(calculated_text)
        else:
            if ach_values:
                latest_ach_value = ach_values[-1]
                if latest_ach_value == 'Overload' or math.isnan(latest_ach_value):
                    ach_text = f"{self.ach_desc}: Overload"
                else:
                    formatted_value = format_si_unit(latest_ach_value, self.ach_unit)
                    ach_text = f"{self.ach_desc}: {formatted_value}"
                self.value_label_ach.setText(ach_text)
            if self.bch_desc and bch_values:
                latest_bch_value = bch_values[-1]
                if latest_bch_value == 'Overload' or math.isnan(latest_bch_value):
                    bch_text = f"{self.bch_desc}: Overload"
                else:
                    formatted_value = format_si_unit(latest_bch_value, self.bch_unit)
                    bch_text = f"{self.bch_desc}: {formatted_value}"
                self.value_label_bch.setText(bch_text)

        if self.recording:
            if self.jig_mode and calculated_values:
                for t, c_val in zip(time_values, calculated_values):
                    if not math.isinf(c_val):
                        self.recorded_data.append([t, c_val])
            else:
                # bch_values or [None]*len(ach_values) -- 安全にzipする
                combined_b = bch_values if bch_values else [None] * len(ach_values)
                for t, a_val, b_val in zip(time_values, ach_values, combined_b):
                    if self.bch_desc:
                        self.recorded_data.append([t, a_val, b_val])
                    else:
                        self.recorded_data.append([t, a_val])

    def toggle_recording(self):
        if not self.recording:
            self.recording = True
            self.recorded_data = []
            self.recording_button.setText("記録停止")
        else:
            self.recording = False
            self.recording_button.setText("記録開始")
            self.save_recorded_data()

    def save_recorded_data(self):
        if not self.recorded_data:
            QMessageBox.information(self, "情報", "記録されたデータがありません。")
            return

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getSaveFileName(
            self, "CSVファイルに保存", "", "CSV Files (*.csv);;All Files (*)", options=options
        )
        if file_path:
            try:
                with open(file_path, mode='w', newline='') as file:
                    writer = csv.writer(file)
                    if self.jig_mode:
                        writer.writerow(["時間 (秒)", f"{self.calculated_value_desc} ({self.calculated_unit})"])
                        writer.writerows(self.recorded_data)
                    else:
                        if self.bch_desc:
                            writer.writerow(["時間 (秒)",
                                             f"{self.ach_desc} ({self.ach_unit})",
                                             f"{self.bch_desc} ({self.bch_unit})"])
                            writer.writerows(self.recorded_data)
                        else:
                            writer.writerow(["時間 (秒)", f"{self.ach_desc} ({self.ach_unit})"])
                            for row in self.recorded_data:
                                writer.writerow(row)
                QMessageBox.information(self, "成功", f"データを{file_path}に保存しました。")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"データの保存に失敗しました。\n{e}")

    def reset_graph(self):
        self.time_input.setText("10")
        self.ach_data.clear()
        self.bch_data.clear()
        self.calculated_data.clear()
        self.time_data.clear()
        self.start_time = time.time()
        self.figure.clf()
        self.set_measurement_mode_descriptions(
            self.ach_desc,
            self.bch_desc,
            self.jig_mode,
            self.calculated_value_desc,
            self.calculated_unit
        )
        self.canvas.draw()
        self.recorded_data.clear()


class DMMApp(QMainWindow):
    def __init__(self):
        super().__init__(parent=None)
        self.is_ready_event = None
        self.setup_timer = None
        self.setup_dialog = None
        self.rm = rm
        self.resources = self.rm.list_resources()
        self.selected_resource = None
        self.trigger_mode = None
        self.measurement_mode_ach = None
        self.measurement_mode_bch = None
        self.ach_unit = ""
        self.bch_unit = ""
        self.jig_mode = False
        self.jig_mode_name = None
        self.calculated_value_desc = ""
        self.calculated_unit = ""

        self.jig_modes = [
            ("四端子抵抗測定A_V", "F1", "F35"),
            ("四端子抵抗測定B_V", "F12", "F5"),
            ("hFE測定", "F5", "F35"),
            ("電力計測(小電流)", "F5", "F12"),
            ("電力計測(大電流)", "F1", "F35"),
        ]

        self.setWindowTitle("7352A/E コントローラー")
        self.setGeometry(100, 100, 1920, 1080)
        font = QFont()
        font.setPointSize(14)
        self.setFont(font)

        self.stacked_widget = QStackedWidget(parent=None)
        self.setCentralWidget(self.stacked_widget)

        self.device_selection_page = DeviceSelectionPage(self.resources)
        self.dmm_setup_page = DMMSetupPage(self.jig_modes)
        self.mode_selection_page = ModeSelectionPage()
        self.value_display_page = ValueDisplayPage()
        self.graph_display_page = GraphDisplayPage()

        self.stacked_widget.addWidget(self.device_selection_page)
        self.stacked_widget.addWidget(self.dmm_setup_page)
        self.stacked_widget.addWidget(self.mode_selection_page)
        self.stacked_widget.addWidget(self.value_display_page)
        self.stacked_widget.addWidget(self.graph_display_page)

        self.device_selection_page.next_button.clicked.connect(self.go_to_dmm_setup)
        self.dmm_setup_page.next_button.clicked.connect(self.go_to_mode_selection)
        self.mode_selection_page.value_display_button.clicked.connect(self.go_to_value_display)
        self.mode_selection_page.graph_display_button.clicked.connect(self.go_to_graph_display)
        self.value_display_page.switch_display_button.clicked.connect(self.switch_display_mode)
        self.graph_display_page.switch_display_button.clicked.connect(self.switch_display_mode)
        self.value_display_page.trigger_button.clicked.connect(self.send_trigger)
        self.graph_display_page.trigger_button.clicked.connect(self.send_trigger)

        self.mode_selection_page.reset_button.clicked.connect(self.reset_application)
        self.value_display_page.reset_button.clicked.connect(self.reset_application)
        self.graph_display_page.reset_button.clicked.connect(self.reset_application)

        self.stacked_widget.setCurrentWidget(self.device_selection_page)

        self.manager = Manager()
        self.command_queue = self.manager.Queue()
        self.data_list = self.manager.list()
        self.measurement_process = None
        self.stop_event = Event()
        self.timer = None
        self.last_read_index = 0

    def switch_display_mode(self):
        self.stacked_widget.setCurrentWidget(self.mode_selection_page)

    def go_to_dmm_setup(self):
        self.selected_resource = self.device_selection_page.combo.currentText()
        if not self.selected_resource:
            QMessageBox.warning(self, "警告", "機器を選択してください。")
            return
        self.stacked_widget.setCurrentWidget(self.dmm_setup_page)

    def go_to_mode_selection(self):
        commands = self.dmm_setup_page.get_setup_commands()
        if not commands:
            return
        self.trigger_mode = self.dmm_setup_page.get_trigger_mode()
        (self.measurement_mode_ach_desc, self.measurement_mode_ach, self.ach_unit), \
            (self.measurement_mode_bch_desc, self.measurement_mode_bch, self.bch_unit) = \
            self.dmm_setup_page.get_measurement_modes()

        mode_id = self.dmm_setup_page.mode_group.checkedId()
        self.jig_mode = (mode_id == 1)
        if self.jig_mode:
            self.jig_mode_name = self.dmm_setup_page.jig_selection_combo.currentText()
            if self.jig_mode_name in ("四端子抵抗測定A_V", "四端子抵抗測定B_V"):
                self.calculated_value_desc = "抵抗値"
                self.calculated_unit = "Ω"
            elif self.jig_mode_name == "hFE測定":
                self.calculated_value_desc = "hFE"
                self.calculated_unit = ""
            elif self.jig_mode_name in ("電力計測(小電流)", "電力計測(大電流)"):
                self.calculated_value_desc = "電力"
                self.calculated_unit = "W"
            else:
                self.calculated_value_desc = ""
                self.calculated_unit = ""
        else:
            self.jig_mode_name = None
            self.calculated_value_desc = ""
            self.calculated_unit = ""

        self.value_display_page.set_measurement_mode_descriptions(
            self.measurement_mode_ach_desc,
            self.measurement_mode_bch_desc,
            self.jig_mode,
            calculated_value_desc=self.calculated_value_desc,
            calculated_unit=self.calculated_unit
        )
        self.graph_display_page.set_measurement_mode_descriptions(
            self.measurement_mode_ach_desc,
            self.measurement_mode_bch_desc,
            self.jig_mode,
            calculated_value_desc=self.calculated_value_desc,
            calculated_unit=self.calculated_unit
        )

        self.start_measurement(commands)

        self.setup_dialog = QDialog(self)
        self.setup_dialog.setWindowTitle("DMMセットアップ中")
        layout = QVBoxLayout()
        label = QLabel("DMMセットアップ中... お待ちください。")
        layout.addWidget(label)
        self.setup_dialog.setLayout(layout)
        self.setup_dialog.setModal(True)

        self.setup_timer = QTimer(self.setup_dialog)
        self.setup_timer.timeout.connect(self.check_dmm_ready)
        self.setup_timer.start(100)

        self.setup_dialog.show()
        self.stacked_widget.setCurrentWidget(self.mode_selection_page)

    def check_dmm_ready(self):
        if self.is_ready_event.is_set():
            self.setup_timer.stop()
            self.setup_dialog.accept()
            self.stacked_widget.setCurrentWidget(self.mode_selection_page)

    def start_measurement(self, setup_commands):
        if self.measurement_process and self.measurement_process.is_alive():
            self.stop_measurement()
        self.stop_event.clear()

        self.is_ready_event = Event()
        self.measurement_process = MeasurementClass(
            self.command_queue,
            self.data_list,
            self.selected_resource,
            self.stop_event,
            self.is_ready_event
        )
        self.measurement_process.start()

        self.command_queue.put("SEND *RST")
        for cmd in setup_commands:
            self.command_queue.put(f"SEND {cmd}")

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_from_shared_memory)
        self.timer.start(20)

    def go_to_value_display(self):
        self.stacked_widget.setCurrentWidget(self.value_display_page)
        if self.trigger_mode == "TRS3":
            self.value_display_page.trigger_button.setEnabled(True)
        else:
            self.value_display_page.trigger_button.setEnabled(False)

    def go_to_graph_display(self):
        self.stacked_widget.setCurrentWidget(self.graph_display_page)
        if self.trigger_mode == "TRS3":
            self.graph_display_page.trigger_button.setEnabled(True)
        else:
            self.graph_display_page.trigger_button.setEnabled(False)

        self.graph_display_page.load_data_from_list(
            self.data_list, self.jig_mode, self.jig_mode_name, self.measurement_mode_bch
        )

    def update_from_shared_memory(self):
        ach_values = []
        bch_values = []
        calculated_values = []
        time_values = []

        new_data_count = len(self.data_list) - self.last_read_index
        if new_data_count <= 0:
            return

        data_slice = self.data_list[self.last_read_index:]
        self.last_read_index += new_data_count

        for timestamp, ach_value, bch_value in data_slice:
            time_values.append(timestamp)

            if self.jig_mode:
                if self.jig_mode_name == "四端子抵抗測定A_V":
                    if ach_value != 0 and not math.isnan(ach_value) and not math.isnan(bch_value):
                        try:
                            calculated_value = abs(ach_value / bch_value)
                        except ZeroDivisionError:
                            calculated_value = float('inf')  # 修正: 割り切れない場合はinf
                    else:
                        calculated_value = float('inf')  # 修正
                elif self.jig_mode_name == "四端子抵抗測定B_V":
                    if ach_value != 0 and not math.isnan(ach_value) and not math.isnan(bch_value):
                        try:
                            calculated_value = abs(ach_value / bch_value)
                        except ZeroDivisionError:
                            calculated_value = float('inf')  # 修正
                    else:
                        calculated_value = float('inf')  # 修正
                elif self.jig_mode_name == "hFE測定":
                    if ach_value != 0 and not math.isnan(ach_value) and not math.isnan(bch_value):
                        try:
                            calculated_value = bch_value / ach_value
                        except ZeroDivisionError:
                            calculated_value = float('inf')  # 修正
                    else:
                        calculated_value = float('inf')  # 修正
                elif self.jig_mode_name == "電力計測(小電流)":
                    if not math.isnan(ach_value) and not math.isnan(bch_value):
                        calculated_value = ach_value * bch_value
                    else:
                        calculated_value = float('inf')  # 修正
                elif self.jig_mode_name == "電力計測(大電流)":
                    if not math.isnan(ach_value) and not math.isnan(bch_value):
                        calculated_value = ach_value * bch_value
                    else:
                        calculated_value = float('inf')  # 修正
                else:
                    calculated_value = float('inf')  # 修正
                calculated_values.append(calculated_value)
            else:
                ach_values.append(ach_value)
                if self.measurement_mode_bch:
                    bch_values.append(bch_value)

        # 値表示
        if self.stacked_widget.currentWidget() == self.value_display_page:
            if self.jig_mode:
                if calculated_values:
                    latest_calculated = calculated_values[-1]
                    self.value_display_page.update_values(
                        ach_value=None,
                        bch_value=None,
                        calculated_value=latest_calculated,
                        ach_unit=self.ach_unit,
                        bch_unit=self.bch_unit,
                        calculated_unit=self.calculated_unit,
                        jig_mode=self.jig_mode
                    )
            else:
                if ach_values:
                    latest_ach = ach_values[-1]
                    latest_bch = bch_values[-1] if bch_values else None
                    self.value_display_page.update_values(
                        ach_value=latest_ach,
                        bch_value=latest_bch,
                        calculated_value=None,
                        ach_unit=self.ach_unit,
                        bch_unit=self.bch_unit,
                        jig_mode=self.jig_mode
                    )

        # グラフ表示
        if self.stacked_widget.currentWidget() == self.graph_display_page:
            if self.jig_mode and calculated_values:
                self.graph_display_page.update_graph(
                    ach_values=[],
                    bch_values=[],
                    time_values=time_values,
                    calculated_values=calculated_values
                )
            elif ach_values:
                self.graph_display_page.update_graph(
                    ach_values=ach_values,
                    bch_values=bch_values if bch_values else None,
                    time_values=time_values
                )

    def stop_measurement(self):
        if self.measurement_process and self.measurement_process.is_alive():
            self.command_queue.put("STOP")
            self.measurement_process.join()
            self.measurement_process = None
        if self.timer and self.timer.isActive():
            self.timer.stop()
        self.stacked_widget.setCurrentWidget(self.mode_selection_page)

    def send_trigger(self):
        if self.trigger_mode == "TRS3":
            self.command_queue.put("TRIGGER")

    def reset_application(self):
        if self.measurement_process and self.measurement_process.is_alive():
            self.command_queue.put("STOP")
            self.measurement_process.join()
            self.measurement_process = None
        self.selected_resource = None
        self.trigger_mode = None
        self.measurement_mode_ach = None
        self.measurement_mode_bch = None
        self.ach_unit = ""
        self.bch_unit = ""
        self.jig_mode = False
        self.jig_mode_name = None
        self.calculated_value_desc = ""
        self.calculated_unit = ""

        self.graph_display_page.reset_graph()
        self.value_display_page.set_measurement_mode_descriptions("", "", self.jig_mode)
        self.stacked_widget.setCurrentWidget(self.device_selection_page)

    def closeEvent(self, event, *args, **kwargs):
        if self.measurement_process and self.measurement_process.is_alive():
            self.command_queue.put("STOP")
            self.measurement_process.join()
        event.accept()


def main():
    app = QApplication(sys.argv)
    font = QFont()
    font.setPointSize(14)
    app.setFont(font, "QLabel")

    window = DMMApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
