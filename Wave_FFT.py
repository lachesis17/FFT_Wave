import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QDoubleSpinBox, QPushButton, QSpinBox, QDesktopWidget
from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import QtGui
import numpy as np
import matplotlib.pyplot as plt
import scipy.fft as fft
import pandas as pd
import ctypes

appid = 'Wave_Analysis'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(appid)
QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Wave Parameters")
        self.setGeometry(50, int(QDesktopWidget().screenGeometry().height() / 2) - 150, 400, 300)
        self.setWindowIcon(QtGui.QIcon('assets/icon.svg'))
        self.setWindowFlags(QtCore.Qt.WindowType.WindowStaysOnTopHint) # Force top level

        self.plot_window = None
        layout = QVBoxLayout()

        self.amplitude_input = QSpinBox()
        self.amplitude_input.setRange(0, 1000)
        self.amplitude_input.setValue(1)
        layout.addWidget(QLabel("Amplitude"))
        layout.addWidget(self.amplitude_input)

        self.frequency1_input = QSpinBox()
        self.frequency1_input.setRange(0, 100000)
        self.frequency1_input.setValue(1000)
        layout.addWidget(QLabel("Frequency 1 (Hz)"))
        layout.addWidget(self.frequency1_input)

        self.frequency2_input = QSpinBox()
        self.frequency2_input.setRange(0, 100000)
        self.frequency2_input.setValue(5000)
        layout.addWidget(QLabel("Frequency 2 (Hz)"))
        layout.addWidget(self.frequency2_input)

        self.sample_rate_input = QSpinBox()
        self.sample_rate_input.setRange(0, 1000000)
        self.sample_rate_input.setValue(44000)
        layout.addWidget(QLabel("Sample Rate (Hz)"))
        layout.addWidget(self.sample_rate_input)

        self.duration_input = QDoubleSpinBox()
        self.duration_input.setRange(0, 0.10)
        self.duration_input.setSingleStep(0.001)
        self.duration_input.setDecimals(3)
        self.duration_input.setValue(0.005)
        layout.addWidget(QLabel("Duration (s)"))
        layout.addWidget(self.duration_input)

        self.plot_button = QPushButton("Plot")
        self.plot_button.clicked.connect(self.plot)
        layout.addWidget(self.plot_button)

        self.setLayout(layout)

    def do_fft(self, duration, y):
        fft_s = fft.fft(y, overwrite_x=False)
        time = duration
        num_samples = len(y)
        sample_rate = num_samples / time
        freq = (sample_rate / num_samples) * np.arange(0, (num_samples / 2) + 1)
        amp = np.abs(fft_s)[0:(np.int_(len(fft_s) / 2) + 1)]

        if len(freq) != len(amp):
            if len(freq) > len(amp):
                c = len(freq) - len(amp)
                freq = freq[:len(freq) - c]
            else:
                c = len(amp) - len(freq)
                amp = amp[:len(amp) - c]

        return freq, amp


    def plot(self):
        amp = self.amplitude_input.value()
        f1 = self.frequency1_input.value()
        f2 = self.frequency2_input.value()
        fs = self.sample_rate_input.value()
        duration = self.duration_input.value()

        print("Amplitude:", amp)
        print("Frequency 1:", f1)
        print("Frequency 2:", f2)
        print("Sample Rate:", fs)
        print("Duration:", duration)

        t1 = np.linspace(0, duration, int(fs * duration)) # Time 1
        t2 = np.linspace(0, duration, int(fs * duration)) # Time 2

        y1 = amp * np.sin(2 * np.pi * f1 * t1) # Wave 1
        y2 = amp * np.sin(2 * np.pi * f2 * t2) # Wave 2

        if self.plot_window is not None:
            plt.close(self.plot_window)

        self.plot_window = plt.figure(figsize=(10, 6)) # Plot window

        df = pd.DataFrame() # Dataframe
        df[f'{f1}khz Freq'] = y1
        df[f'{f1}khz Time'] = t1
        df[f'{f2}khz Freq'] = y2
        df[f'{f2}khz Time'] = t2
        df['Sample Rate'] = fs

        plt.subplot(3, 2, 1)
        plt.plot(t1, y1)
        plt.title(f'Sine Wave at {f1}kHz', fontsize=8)
        plt.xlabel('Time (s)', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        plt.subplot(3, 2, 2)
        plt.plot(t2, y2)
        plt.title(f'Sine Wave at {f2}kHz', fontsize=8)
        plt.xlabel('Time (s)', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        freq, amp = self.do_fft(y=y1, duration=duration)

        df2 = pd.DataFrame() # Dataframe
        df2[f'{f1}khz FFT Freq'] = freq
        df2[f'{f1}khz FFT Amp'] = amp

        idx_max_first = np.argmax(amp)
        print("Index of Max Amplitude: ", idx_max_first)
        print("Max Amp of 1kHz: ", amp[idx_max_first], "Frequency of max Amp of 1kHz: ", freq[idx_max_first])

        df3 = pd.DataFrame()
        df3['Max Amp (First Wave)'] = [amp[idx_max_first]]
        df3['Freq of Max Amp (First Wave)'] = [freq[idx_max_first]]

        freq = freq[:idx_max_first*2+1]
        amp = amp[:idx_max_first*2+1]

        plt.subplot(3, 2, 3)
        plt.plot(freq, amp)
        plt.title(f'{f1}kHz (FFT)', fontsize=8)
        plt.xlabel('Frequency', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        freq, amp = self.do_fft(y=y2, duration=duration)

        df2[f'{f2}khz FFT Freq'] = freq
        df2[f'{f2}khz FFT Amp'] = amp

        idx_max_second = np.argmax(amp)
        print("Index of Max Amplitude: ", idx_max_second)
        print("Max Amp of 5kHz: ", amp[idx_max_second], "Frequency of max Amp of 5kHz: ", freq[idx_max_second])

        df3['Max Amp (Second Wave)'] = [amp[idx_max_second]]
        df3['Freq of Max Amp (Second Wave)'] = [freq[idx_max_second]]

        freq = freq[:idx_max_second*2+1]
        amp = amp[:idx_max_second*2+1]

        plt.subplot(3, 2, 4)
        plt.plot(freq, amp)
        plt.title(f'{f2}kHz (FFT)', fontsize=8)
        plt.xlabel('Frequency', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        # Add both sine waves
        y_combined = y1 + y2

        plt.subplot(3, 2, 5)
        plt.plot(t1, y_combined)
        plt.title('Combined Waves', fontsize=8)
        plt.xlabel('Time (s)', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        freq, amp = self.do_fft(y=y_combined, duration=duration)

        df2['Combined FFT Freq'] = freq
        df2['Combined FFT Amp'] = amp

        plt.subplot(3, 2, 6)
        plt.plot(freq, amp)
        plt.title('Combined Waves (FFT)', fontsize=8)
        plt.xlabel('Frequency', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)
        plt.xlim(0, f2*2)

        df2 = pd.concat([df2, df3], axis=1)
        df = pd.concat([df, df2], axis=1)
        df.to_excel("fft.xlsx")

        plt.tight_layout()
        plt.show()

    def closeEvent(self, a0: QtGui.QCloseEvent) -> None:
        if self.plot_window is not None:
            plt.close(self.plot_window)

        return super().closeEvent(a0)

def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    sys.excepthook = except_hook
    main()


