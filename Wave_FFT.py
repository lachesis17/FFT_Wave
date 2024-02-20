import sys
from PyQt5.QtWidgets import QApplication, QDesktopWidget, QAbstractSpinBox, QMainWindow
from PyQt5 import QtCore, QtWidgets, QtGui, uic
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

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("assets/MainWindow.ui", self)
        self.setWindowIcon(QtGui.QIcon('assets/icon.svg'))
        self.setWindowFlags(QtCore.Qt.WindowType.WindowStaysOnTopHint) # Force top level
        self.installEventFilter(self)
        self.move_to_primary_monitor()

        self.plot_window = None
        self.media_player = QMediaPlayer()

        self.play_wave_1_but.clicked.connect(self.play_wave_1)
        self.play_wave_2_but.clicked.connect(self.play_wave_2)
        self.play_combined_but.clicked.connect(self.play_combined)
        self.plot_button.clicked.connect(self.plot)

        self.frequency1_input.setStepType(QAbstractSpinBox.AdaptiveDecimalStepType)
        self.frequency1_input.valueChanged.connect(self.change_default_duration)
        self.frequency2_input.setStepType(QAbstractSpinBox.AdaptiveDecimalStepType)
        self.frequency2_input.valueChanged.connect(self.change_default_duration)
        self.sample_rate_input.setStepType(QAbstractSpinBox.AdaptiveDecimalStepType)


    def play_wave_1(self):
        self.media_player.stop()
        self.media_player.setMedia(QMediaContent(None))

        samplerate = self.sample_rate_input.value()
        t = np.linspace(0., 1., samplerate)
        amplitude = np.iinfo(np.int16).max
        fs = self.frequency1_input.value()
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

        samplerate = self.sample_rate_input.value()
        t = np.linspace(0., 1., samplerate)
        amplitude = np.iinfo(np.int16).max
        fs = self.frequency2_input.value()
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

        samplerate = self.sample_rate_input.value()
        t = np.linspace(0., 1., samplerate)
        amplitude = np.iinfo(np.int16).max
        fs1 = self.frequency1_input.value()
        data1 = amplitude * np.sin(2. * np.pi * fs1 * t)
        fs2 = self.frequency2_input.value()
        data2 = amplitude * np.sin(2. * np.pi * fs2 * t)
        data = data1 + data2
        wave_file = "Wave_Combined.wav"
        write_wav(wave_file, samplerate, data.astype(np.int16))

        media = QMediaContent(QUrl.fromLocalFile(wave_file))
        self.media_player.setMedia(media)
        self.media_player.setVolume(20)
        self.media_player.play()

    def change_default_duration(self):
        # when the min frequency is too small without a long enough duration, the increments for FFT don't give correct indexes
        max_freq = max([self.frequency1_input.value(), self.frequency2_input.value()])
        min_freq = min([self.frequency1_input.value(), self.frequency2_input.value()])
        if min_freq < 100:
            self.duration_input.setValue(0.035)
            return
        if min_freq <= 200:
            self.duration_input.setValue(0.020)
            return
        if max_freq < 10000:
            self.duration_input.setValue(0.005)
            return
        if max_freq > 10000 and min_freq > 10000:
            self.duration_input.setValue(0.002)
            return


    def do_fft(self, duration, y):
        fft_amp = fft.fft(y, overwrite_x=False)
        num_samples = len(y)
        sample_rate = num_samples / duration
        freq = (sample_rate / num_samples) * np.arange(0, (num_samples / 2) + 1)
        amp = np.abs(fft_amp)[0:(np.int_(len(fft_amp) / 2) + 1)]

        if len(freq) != len(amp):
            if len(freq) > len(amp):
                diff = len(freq) - len(amp)
                freq = freq[:len(freq) - diff]
            else:
                diff = len(amp) - len(freq)
                amp = amp[:len(amp) - diff]

        return freq, amp


    def plot(self):
        if self.plot_window is not None:
            plt.close(self.plot_window)
    
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
        self.plot_window.show()

    def move_to_primary_monitor(self):
        desktop = QApplication.desktop()
        width = 0
        active_screen = None

        for screen in range(desktop.screenCount()):
            screen_geometry = desktop.screenGeometry(screen)
            if screen_geometry.width() > width:
                width = screen_geometry.width()
                active_screen = screen

        if active_screen is not None:
            screen_geometry = desktop.screenGeometry(active_screen)
            self.move(screen_geometry.topLeft())
            self.setGeometry(int((QDesktopWidget().screenGeometry().width() / 2) * 0.05), int(QDesktopWidget().screenGeometry().height() / 2) - 150, 400, 200)

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


