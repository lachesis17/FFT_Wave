import sys
from PyQt5.QtWidgets import QApplication, QDesktopWidget, QAbstractSpinBox, QMainWindow, QButtonGroup
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

        self.wave_1_group = QButtonGroup()
        wave_1_types = [self.sine_1, self.square_1, self.triangle_1, self.saw_1]
        for wave in wave_1_types:
            self.wave_1_group.addButton(wave)
        self.sine_1.setChecked(True)

        self.wave_2_group = QButtonGroup()
        wave_2_types = [self.sine_2, self.square_2, self.triangle_2, self.saw_2]
        for wave in wave_2_types:
            self.wave_2_group.addButton(wave)
        self.sine_2.setChecked(True)

        self.frequency1_input.setStepType(QAbstractSpinBox.AdaptiveDecimalStepType)
        self.frequency1_input.valueChanged.connect(self.change_default_duration)
        self.frequency2_input.setStepType(QAbstractSpinBox.AdaptiveDecimalStepType)
        self.frequency2_input.valueChanged.connect(self.change_default_duration)
        self.sample_rate_input.setStepType(QAbstractSpinBox.AdaptiveDecimalStepType)

    def sine_wave(self, amp, fs, t):
        return amp * np.sin(2. * np.pi * fs * t)
    
    def square_wave(self, amp, fs, t):
        return amp * np.sign(np.sin(2. * np.pi * fs * t))

    def triangle_wave(self, amp, fs, t):
        return amp * (2. / np.pi) * np.arcsin(np.sin(2. * np.pi * fs * t))

    def saw_wave(self, amp, fs, t):
        return amp * (2. * (fs * t - np.floor(0.5 + fs * t)))

    def play_wave_1(self):
        self.media_player.stop()
        self.media_player.setMedia(QMediaContent(None))

        samplerate = self.sample_rate_input.value()
        t = np.linspace(0., 1., samplerate)
        amplitude = np.iinfo(np.int16).max
        fs = self.frequency1_input.value()

        if self.sine_1.isChecked():
            data = self.sine_wave(amp=amplitude, fs=fs, t=t) # Wave 1 Sine
        elif self.square_1.isChecked():
            data = self.square_wave(amp=amplitude, fs=fs, t=t) # Wave 1 Square
        elif self.triangle_1.isChecked():
            data = self.triangle_wave(amp=amplitude, fs=fs, t=t) # Wave 1 Triangle
        elif self.saw_1.isChecked():
            data = self.saw_wave(amp=amplitude, fs=fs, t=t) # Wave 1 Saw
        
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

        if self.sine_2.isChecked():
            data = self.sine_wave(amp=amplitude, fs=fs, t=t) # Wave 2 Sine
        elif self.square_2.isChecked():
            data = self.square_wave(amp=amplitude, fs=fs, t=t) # Wave 2 Square
        elif self.triangle_2.isChecked():
            data = self.triangle_wave(amp=amplitude, fs=fs, t=t) # Wave 2 Triangle
        elif self.saw_2.isChecked():
            data = self.saw_wave(amp=amplitude, fs=fs, t=t) # Wave 2 Saw

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
        if self.sine_1.isChecked():
            data1 = self.sine_wave(amp=amplitude, fs=fs1, t=t) # Wave 1 Sine
        elif self.square_1.isChecked():
            data1 = self.square_wave(amp=amplitude, fs=fs1, t=t) # Wave 1 Square
        elif self.triangle_1.isChecked():
            data1 = self.triangle_wave(amp=amplitude, fs=fs1, t=t) # Wave 1 Triangle
        elif self.saw_1.isChecked():
            data1 = self.saw_wave(amp=amplitude, fs=fs1, t=t) # Wave 1 Saw

        fs2 = self.frequency2_input.value()
        if self.sine_2.isChecked():
            data2 = self.sine_wave(amp=amplitude, fs=fs2, t=t) # Wave 2 Sine
        elif self.square_2.isChecked():
            data2 = self.square_wave(amp=amplitude, fs=fs2, t=t) # Wave 2 Square
        elif self.triangle_2.isChecked():
            data2 = self.triangle_wave(amp=amplitude, fs=fs2, t=t) # Wave 2 Triangle
        elif self.saw_2.isChecked():
            data2 = self.saw_wave(amp=amplitude, fs=fs2, t=t) # Wave 2 Saw

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


    def apply_fft(self, duration, y):
        num_samples = len(y)
        if self.hann_check.isChecked():
            y = y * np.hanning(num_samples)   # hann window = 0.5 * (1 - np.cos(2 * np.pi * np.arange(num_samples) / (num_samples - 1)))
        
        fft_mode = self.fft_combo.currentText()
        print("FFT Mode: ", fft_mode)
        
        if fft_mode == "FFT - Fast Fourier Transform":
            fft_amp = fft.fft(y, overwrite_x=False)
        elif fft_mode == "IFFT - Inverse FFT":
            fft_amp = fft.ifft(y, overwrite_x=False)
        elif fft_mode == "RFFT- FFT of strictly real-valued sequence":
            fft_amp = fft.rfft(y, overwrite_x=False)
        elif fft_mode == "IRFFT - Inverse of RFFT":
            fft_amp = fft.irfft(y, overwrite_x=False)
        elif fft_mode == "HFFT - FFT of a Hermitian sequence (real spectrum)":
            fft_amp = fft.hfft(y, overwrite_x=False)
        elif fft_mode == "IHFFT - Inverse of HFFT":
            fft_amp = fft.ihfft(y, overwrite_x=False)
        elif fft_mode == "DCT - Discrete cosine transform":
            fft_amp = fft.dct(y, overwrite_x=False)
        elif fft_mode == "IDCT - Inverse DCT":
            fft_amp = fft.idct(y, overwrite_x=False)
        elif fft_mode == "DST - Discrete sine transform":
            fft_amp = fft.dst(y, overwrite_x=False)
        elif fft_mode == "IDST - Inverse DST":
            fft_amp = fft.idst(y, overwrite_x=False)

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
        fs1 = self.frequency1_input.value()
        fs2 = self.frequency2_input.value()
        f0 = self.sample_rate_input.value()
        duration = self.duration_input.value()

        print("Amplitude:", amp)
        print("Frequency 1:", fs1)
        print("Frequency 2:", fs2)
        print("Sample Rate:", f0)
        print("Duration:", duration)

        t = np.linspace(0, duration, int(f0 * duration)) # Time

        if self.sine_1.isChecked():
            wave_1_type = "Sine"
            y1 = self.sine_wave(amp=amp, fs=fs1, t=t) # Wave 1 Sine
        elif self.square_1.isChecked():
            wave_1_type = "Square"
            y1 = self.square_wave(amp=amp, fs=fs1, t=t) # Wave 1 Square
        elif self.triangle_1.isChecked():
            wave_1_type = "Triangle"
            y1 = self.triangle_wave(amp=amp, fs=fs1, t=t) # Wave 1 Triangle
        elif self.saw_1.isChecked():
            wave_1_type = "Saw"
            y1 = self.saw_wave(amp=amp, fs=fs1, t=t) # Wave 1 Saw
        
        if self.sine_2.isChecked():
            wave_2_type = "Sine"
            y2 = self.sine_wave(amp=amp, fs=fs2, t=t) # Wave 2 Sine
        elif self.square_2.isChecked():
            wave_2_type = "Square"
            y2 = self.square_wave(amp=amp, fs=fs2, t=t) # Wave 2 Square
        elif self.triangle_2.isChecked():
            wave_2_type = "Triangle"
            y2 = self.triangle_wave(amp=amp, fs=fs2, t=t) # Wave 2 Triangle
        elif self.saw_2.isChecked():
            wave_2_type = "Saw"
            y2 = self.saw_wave(amp=amp, fs=fs2, t=t) # Wave 2 Saw

        self.plot_window = plt.figure(num="Wave Frequency Spectrums", figsize=(10, 6)) # Plot window

        df = pd.DataFrame()
        df[f'{fs1} hz Freq'] = y1
        df[f'{fs1} hz Time'] = t
        df[f'{fs2} hz Freq'] = y2
        df[f'{fs2} hz Time'] = t

        plt.subplot(3, 2, 1)
        plt.plot(t, y1)
        plt.title(f'{wave_1_type} Wave at {fs1} Hz', fontsize=8)
        plt.xlabel('Time (s)', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        plt.subplot(3, 2, 2)
        plt.plot(t, y2)
        plt.title(f'{wave_2_type} Wave at {fs2} Hz', fontsize=8)
        plt.xlabel('Time (s)', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        freq, amp = self.apply_fft(y=y1, duration=duration)

        df2 = pd.DataFrame()
        df2[f'FFT - {fs1} hz Freq'] = freq
        df2[f'FFT - {fs1} hz FFT Amp'] = amp

        idx_max_first = np.argmax(amp)
        print("Index of Max Amplitude Wave 1: ", idx_max_first)
        print(f"Max Amp of {fs1} Hz: ", amp[idx_max_first], f"Frequency of max Amp of {fs1} Hz: ", freq[idx_max_first])

        df3 = pd.DataFrame()
        df3['Sample Rate'] = [f0]
        df3['Max Amp (First Wave)'] = [amp[idx_max_first]]
        df3['Freq of Max Amp (First Wave)'] = [freq[idx_max_first]]

        if self.scale_axes.isChecked():
            freq = freq[:idx_max_first*2+1]
            amp = amp[:idx_max_first*2+1]

        plt.subplot(3, 2, 3)
        plt.plot(freq, amp)
        plt.title(f'{fs1} Hz (FFT)', fontsize=8)
        plt.xlabel('Frequency', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        freq, amp = self.apply_fft(y=y2, duration=duration)

        df2[f'FFT - {fs2} hz Freq'] = freq
        df2[f'FFT - {fs2} hz FFT Amp'] = amp

        idx_max_second = np.argmax(amp)
        print("Index of Max Amplitude Wave 2: ", idx_max_second)
        print(f"Max Amp of {fs2} Hz: ", amp[idx_max_second], f"Frequency of max Amp of {fs2} Hz: ", freq[idx_max_second])

        df3['Max Amp (Second Wave)'] = [amp[idx_max_second]]
        df3['Freq of Max Amp (Second Wave)'] = [freq[idx_max_second]]

        if self.scale_axes.isChecked():
            freq = freq[:idx_max_second*2+1]
            amp = amp[:idx_max_second*2+1]

        plt.subplot(3, 2, 4)
        plt.plot(freq, amp)
        plt.title(f'{fs2} Hz (FFT)', fontsize=8)
        plt.xlabel('Frequency', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        # Add both sine waves
        y_combined = y1 + y2

        df[f'Combined Freq'] = y_combined
        df[f'Combined Time'] = t

        plt.subplot(3, 2, 5)
        plt.plot(t, y_combined)
        plt.title('Combined Waves', fontsize=8)
        plt.xlabel('Time (s)', fontsize=8)
        plt.ylabel('Amplitude', fontsize=8)
        plt.grid(True)

        freq, amp = self.apply_fft(y=y_combined, duration=duration)

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
        plt.xlim(0, fs2*2)

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
            self.setGeometry(int((QDesktopWidget().screenGeometry().width() / 2) * 0.05), int(QDesktopWidget().screenGeometry().height() / 2) - 238, 400, 200)

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


