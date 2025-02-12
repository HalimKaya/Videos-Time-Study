import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
                             QLineEdit, QTableWidget, QTableWidgetItem, QFileDialog, QSplitter)
from PyQt5.QtCore import Qt
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtMultimediaWidgets import QVideoWidget
from PyQt5.QtCore import QUrl
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows  # dataframe_to_rows import edildi
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import pandas as pd

class VideoAnnotator(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Video Time Window")
        self.setGeometry(100, 100, 1800, 800)  # Genişlik ve yükseklik artırıldı

        # Sol panel: Video oynatma ve kontrol butonları
        self.video_player = QMediaPlayer(None, QMediaPlayer.VideoSurface)
        video_widget = QVideoWidget()
        
        self.video_player.setVideoOutput(video_widget)

        self.play_button = QPushButton("Play")
        self.play_button.clicked.connect(self.play_video)

        self.pause_button = QPushButton("Pause")
        self.pause_button.clicked.connect(self.pause_video)

        self.forward_button = QPushButton("Forward 5s")
        self.forward_button.clicked.connect(self.forward_video)

        self.backward_button = QPushButton("Backward 5s")
        self.backward_button.clicked.connect(self.backward_video)

        self.open_button = QPushButton("Open Video")
        self.open_button.clicked.connect(self.open_file)

        # Sol tarafta video oynatıcı ve butonlar olacak
        video_controls_layout = QHBoxLayout()
        video_controls_layout.addWidget(self.play_button)
        video_controls_layout.addWidget(self.pause_button)
        video_controls_layout.addWidget(self.forward_button)
        video_controls_layout.addWidget(self.backward_button)

        video_layout = QVBoxLayout()
        video_layout.addWidget(video_widget)
        video_layout.addWidget(self.open_button)
        video_layout.addLayout(video_controls_layout)

        # Sağ panel: Point, Save butonları, Text Girişi ve Tablo
        self.text_input = QLineEdit(self)
        self.point_button = QPushButton("Point")
        self.point_button.clicked.connect(self.point_action)

        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self.save_action)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Name", "Start", "End", "Duration"])

        # Pie chart göstermek için QLabel ve FigureCanvas
        self.canvas = FigureCanvas(plt.Figure())  # FigureCanvas oluşturuluyor
        self.chart_label = QLabel(self)
        self.chart_label.setFixedHeight(300)  # Sabit yükseklik veriyoruz
        self.chart_label.setAlignment(Qt.AlignCenter)  # Ortala

        # Sağ tarafta butonlar, tablo ve pie chart olacak
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("Text Input:"))
        right_layout.addWidget(self.text_input)
        right_layout.addWidget(self.point_button)
        right_layout.addWidget(self.save_button)
        right_layout.addWidget(self.table)
        right_layout.addWidget(self.canvas)  # Pie chart alanı

        # Splitter: Sol ve sağ panelleri bölecek
        splitter = QSplitter(Qt.Horizontal)
        left_widget = QWidget()
        left_widget.setLayout(video_layout)
        right_widget = QWidget()
        right_widget.setLayout(right_layout)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)

        # Sol panelin genişliğini daha fazla yap (örneğin 70% sol, 30% sağ)
        splitter.setSizes([700, 300])

        # Ana düzen: Splitter'ı ana pencereye ekle
        main_layout = QVBoxLayout()
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        # Veri depolama
        self.annotations = {}

    def open_file(self):
        # Dosya açma fonksiyonu
        filename, _ = QFileDialog.getOpenFileName(self, "Open Video")
        if filename != '':
            self.video_player.setMedia(QMediaContent(QUrl.fromLocalFile(filename)))
            self.play_button.setEnabled(True)

    def play_video(self):
        self.video_player.play()

    def pause_video(self):
        self.video_player.pause()

    def forward_video(self):
        # 5 saniye ileri sarma
        current_time = self.video_player.position()
        self.video_player.setPosition(current_time + 5000)

    def backward_video(self):
        # 5 saniye geri sarma
        current_time = self.video_player.position()
        self.video_player.setPosition(current_time - 5000)

    def point_action(self):
        # Point butonuna tıklanınca metni ve zaman damgasını kaydet
        current_text = self.text_input.text()
        current_time = self.video_player.position() / 1000  # saniye olarak zaman

        if current_text in self.annotations:
            # Eğer plaka mevcutsa "End" kolonunu güncelle
            row = self.annotations[current_text]
            self.table.setItem(row, 2, QTableWidgetItem(f"{current_time:.2f}"))
            # Süreyi hesapla
            start_time = float(self.table.item(row, 1).text())
            duration = current_time - start_time
            self.table.setItem(row, 3, QTableWidgetItem(f"{duration:.2f}"))
        else:
            # Eğer plaka mevcut değilse yeni satır ekle
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            self.table.setItem(row_position, 0, QTableWidgetItem(current_text))
            self.table.setItem(row_position, 1, QTableWidgetItem(f"{current_time:.2f}"))
            self.annotations[current_text] = row_position

    def save_action(self):
        # Verileri tabloya kaydet ve pie chart oluştur
        row_count = self.table.rowCount()
        if row_count == 0:
            print("Tablo boş!")
            return
        
        data = []
        for row in range(row_count):
            name = self.table.item(row, 0).text()
            duration = float(self.table.item(row, 3).text())
            data.append((name, duration))

        # Verileri pandas DataFrame'e dönüştür
        df = pd.DataFrame(data, columns=["Name", "Duration"])

        # Excel dosyasına yaz
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Annotations"

        # DataFrame'i satır satır Excel'e yaz
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Pie chart oluştur
        labels = df["Name"]
        sizes = df["Duration"]

        self.canvas.figure.clear()  # Önceki grafiği temizle
        ax = self.canvas.figure.add_subplot(111)  # Yeni bir subplot ekle
        ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
        ax.axis('equal')  # Eşit oran

        self.canvas.draw()  # Canvas'ı güncelle

        # Pie chart'ı kaydet
        pie_image_path = "pie_chart.png"
        self.canvas.figure.savefig(pie_image_path)

        # Excel dosyasına pie chart ekle
        img = openpyxl.drawing.image.Image(pie_image_path)
        ws.add_image(img, "F1")

        # Excel dosyasını kaydet
        wb.save("çıktı.xlsx")
        print("Excel dosyası ve Pie Chart kaydedildi!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VideoAnnotator()
    window.show()
    sys.exit(app.exec_())
