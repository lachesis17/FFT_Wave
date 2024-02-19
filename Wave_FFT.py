import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QDoubleSpinBox, QPushButton, QSpinBox, QDesktopWidget, QCheckBox, QAbstractSpinBox
from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5 import QtGui
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtCore import QUrl
from scipy.io.wavfile import write as write_wav
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
        self.media_player = QMediaPlayer()
        layout = QVBoxLayout()

        button_layout = QHBoxLayout()
        self.play_wave_1_but = QPushButton("Play Wave 1")
        self.play_wave_2_but = QPushButton("Play Wave 2")
        self.play_combined_but = QPushButton("Play Combined")
        self.play_wave_1_but.clicked.connect(self.play_wave_1)
        self.play_wave_2_but.clicked.connect(self.play_wave_2)
        self.play_combined_but.clicked.connect(self.play_combined)

        button_layout.addWidget(self.play_wave_1_but)
        button_layout.addWidget(self.play_wave_2_but)
        button_layout.addWidget(self.play_combined_but)
        layout.addLayout(button_layout)

        self.amplitude_input = QSpinBox()
        self.amplitude_input.setRange(0, 1000)
        self.amplitude_input.setValue(1)
        layout.addWidget(QLabel("Amplitude"))
        layout.addWidget(self.amplitude_input)

        self.frequency1_input = QSpinBox()
        self.frequency1_input.setRange(0, 100000)
        self.frequency1_input.setValue(1000)
        self.frequency1_input.setSingleStep(100)
        self.frequency1_input.setStepType(QAbstractSpinBox.AdaptiveDecimalStepType)
        self.frequency1_input.valueChanged.connect(self.change_default_duration)
        layout.addWidget(QLabel("Wave 1 Frequency (Hz)"))
        layout.addWidget(self.frequency1_input)

        self.frequency2_input = QSpinBox()
        self.frequency2_input.setRange(0, 100000)
        self.frequency2_input.setValue(5000)
        self.frequency2_input.setSingleStep(100)
        self.frequency2_input.setStepType(QAbstractSpinBox.AdaptiveDecimalStepType)
        self.frequency2_input.valueChanged.connect(self.change_default_duration)
        layout.addWidget(QLabel("Wave 2 Frequency (Hz)"))
        layout.addWidget(self.frequency2_input)

        self.sample_rate_input = QSpinBox()
        self.sample_rate_input.setRange(0, 1000000)
        self.sample_rate_input.setValue(44000)
        self.sample_rate_input.setSingleStep(100)
        self.sample_rate_input.setStepType(QAbstractSpinBox.AdaptiveDecimalStepType)
        layout.addWidget(QLabel("Sample Rate (Hz)"))
        layout.addWidget(self.sample_rate_input)

        self.duration_input = QDoubleSpinBox()
        self.duration_input.setRange(0.002, 0.20)
        self.duration_input.setSingleStep(0.001)
        self.duration_input.setDecimals(3)
        self.duration_input.setValue(0.020)
        layout.addWidget(QLabel("Duration (s)"))
        layout.addWidget(self.duration_input)

        self.scale_axes = QCheckBox("Scale Axes to Peaks")
        self.scale_axes.setChecked(True)
        layout.addWidget(self.scale_axes)

        self.plot_button = QPushButton("Plot")
        self.plot_button.clicked.connect(self.plot)
        layout.addWidget(self.plot_button)

        self.setLayout(layout)

    def play_wave_1(self):
        self.media_player.stop()
        self.media_player.setMedia(QMediaContent(None))

        samplerate = self.sample_rate_input.value(); fs = self.frequency1_input.value()
        t = np.linspace(0., 1., samplerate)
        amplitude = np.iinfo(np.int16).max
        data = amplitude * np.sin(2. * np.pi * fs * t)
        wave_file = "Wave_1.wav"
        write_wav(wave_file, samplerate, data.astype(np.int16))

        media = QMediaContent(QUrl.fromLocalFile(wave_file))
        self.media_player.setMedia(media)
        self.media_player.setVolume(20)
        self.media_player.play()

    def play_wave_2(self):
        self.media_player.stop()
        self.media_player.setMedia(QMediaContent(None))

        samplerate = self.sample_rate_input.value(); fs = self.frequency2_input.value()
        t = np.linspace(0., 1., samplerate)
        amplitude = np.iinfo(np.int16).max
        data = amplitude * np.sin(2. * np.pi * fs * t)
        wave_file = "Wave_2.wav"
        write_wav(wave_file, samplerate, data.astype(np.int16))

        media = QMediaContent(QUrl.fromLocalFile(wave_file))
        self.media_player.setMedia(media)
        self.media_player.setVolume(20)
        self.media_player.play()

    def play_combined(self):
        self.media_player.stop()
        self.media_player.setMedia(QMediaContent(None))

        samplerate = self.sample_rate_input.value(); fs = self.frequency1_input.value()
        t = np.linspace(0., 1., samplerate)
        amplitude = np.iinfo(np.int16).max
        data1 = amplitude * np.sin(2. * np.pi * fs * t)
        samplerate = self.sample_rate_input.value(); fs = self.frequency2_input.value()
        t = np.linspace(0., 1., samplerate)
        amplitude = np.iinfo(np.int16).max
        data2 = amplitude * np.sin(2. * np.pi * fs * t)
        data = data1 + data2
        wave_file = "Wave_Combined.wav"
        write_wav(wave_file, samplerate, data.astype(np.int16))

        media = QMediaContent(QUrl.fromLocalFile(wave_file))
        self.media_player.setMedia(media)
        self.media_player.setVolume(20)
        self.media_player.play()

    def change_default_duration(self):
        # when the min frequency is too small without a long enough duration, the increments for FFQ don't give correct indexes
        max_freq = max([self.frequency1_input.value(), self.frequency2_input.value()])
        min_freq = min([self.frequency1_input.value(), self.frequency2_input.value()])
        if min_freq < 100:
            self.duration_input.setValue(0.035)
            return
        if min_freq <= 200:
            self.duration_input.setValue(0.020)
            return
        if max_freq < 10000 and max_freq > 1000:
            self.duration_input.setValue(0.005)
            return
        if max_freq > 10000:
            self.duration_input.setValue(0.002)
            return


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

        self.plot_window = plt.figure(num="Sine Wave Frequency Spectrums", figsize=(10, 6)) # Plot window

        df = pd.DataFrame()
        df[f'{f1} hz Freq'] = y1
        df[f'{f1} hz Time'] = t1
        df[f'{f2} hz Freq'] = y2
        df[f'{f2} hz Time'] = t2

        plt.subplot(3, 2, 1)
        plt.plot(t1, y1)
        plt.title(f'Sine Wave at {f1} Hz', fontsize=8)
        plt.xlabel('Time (s)', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        plt.subplot(3, 2, 2)
        plt.plot(t2, y2)
        plt.title(f'Sine Wave at {f2} Hz', fontsize=8)
        plt.xlabel('Time (s)', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        freq, amp = self.do_fft(y=y1, duration=duration)

        df2 = pd.DataFrame()
        df2[f'FFT - {f1} hz Freq'] = freq
        df2[f'FFT - {f1} hz FFT Amp'] = amp

        idx_max_first = np.argmax(amp)
        print("Index of Max Amplitude Wave 1: ", idx_max_first)
        print(f"Max Amp of {f1} Hz: ", amp[idx_max_first], f"Frequency of max Amp of {f1} Hz: ", freq[idx_max_first])

        df3 = pd.DataFrame()
        df3['Sample Rate'] = [fs]
        df3['Max Amp (First Wave)'] = [amp[idx_max_first]]
        df3['Freq of Max Amp (First Wave)'] = [freq[idx_max_first]]

        if self.scale_axes.isChecked():
            freq = freq[:idx_max_first*2+1]
            amp = amp[:idx_max_first*2+1]

        plt.subplot(3, 2, 3)
        plt.plot(freq, amp)
        plt.title(f'{f1} Hz (FFT)', fontsize=8)
        plt.xlabel('Frequency', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        freq, amp = self.do_fft(y=y2, duration=duration)

        df2[f'FFT - {f2} hz Freq'] = freq
        df2[f'FFT - {f2} hz FFT Amp'] = amp

        idx_max_second = np.argmax(amp)
        print("Index of Max Amplitude Wave 2: ", idx_max_second)
        print(f"Max Amp of {f2} Hz: ", amp[idx_max_second], f"Frequency of max Amp of {f2} Hz: ", freq[idx_max_second])

        df3['Max Amp (Second Wave)'] = [amp[idx_max_second]]
        df3['Freq of Max Amp (Second Wave)'] = [freq[idx_max_second]]

        if self.scale_axes.isChecked():
            freq = freq[:idx_max_second*2+1]
            amp = amp[:idx_max_second*2+1]

        plt.subplot(3, 2, 4)
        plt.plot(freq, amp)
        plt.title(f'{f2} Hz (FFT)', fontsize=8)
        plt.xlabel('Frequency', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        # Add both sine waves
        y_combined = y1 + y2

        df[f'Combined Freq'] = y_combined
        df[f'Combined Time'] = t1

        plt.subplot(3, 2, 5)
        plt.plot(t1, y_combined)
        plt.title('Combined Waves', fontsize=8)
        plt.xlabel('Time (s)', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        freq, amp = self.do_fft(y=y_combined, duration=duration)

        df2['FFT - Combined Freq'] = freq
        df2['FFT - Combined FFT Amp'] = amp

        idx_max_combined = np.argpartition(amp, -2)[-2:]
        print("Index of Max Amplitude Combined: ", idx_max_combined)
        print(f"Max Amp of Combined: ", amp[idx_max_combined], f"Frequency of max Amp of Combined: ", freq[idx_max_combined])

        df3['Max Amp (Combined)'] = [amp[idx_max_combined]]
        df3['Freq of Max Amp (Combined)'] = [freq[idx_max_combined]]

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


