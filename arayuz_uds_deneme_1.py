import ctypes
from PCAN_UDS_2013 import *
from ctypes import c_ubyte, c_uint32, c_uint16, byref
import sys, os
from PyQt5.QtCore import QThread, pyqtSignal,   Qt
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QTextEdit,
    QMessageBox,
    QMainWindow,

)

import pandas as pd  # Excel kaydetmek için

# Sabit tipleri tanımlama (ctypes ile)
TPCANHandle = c_uint16  # 16-bit unsigned integer (kanal numarası için)
TPCANBaudrate = c_uint32  # 32-bit unsigned integer (baud rate için)

# PCAN-Basic API'nin kanal ve baud rate sabitleri
PCAN_USBBUS1 = TPCANHandle(0x51)  # PCAN-USB interface, channel 1
PCAN_BAUD_250K = TPCANBaudrate(0x011C)  # Baud rate: 250Kbps (0x011C)

# PCANBasic.dll dosyasının yolu
pcan_basic = ctypes.windll.LoadLibrary(
    'C:\\Users\\kayai\\OneDrive\\Desktop\\visual_Studio_codes\\PCANBasic.dll'
)
pcan_uds = r'C:\\Users\\kayai\\OneDrive\\Desktop\\visual_Studio_codes\\PCAN-UDS.dll'

# TPCANMsg yapısı (CAN mesaj yapısı)
class TPCANMsg(ctypes.Structure):
    _fields_ = [
        ("ID", c_uint32),
        ("MSGTYPE", c_ubyte),
        ("LEN", c_ubyte),
        ("DATA", c_ubyte * 8),
    ]  # 8 byte veri alanı

# CAN hattına bağlanma fonksiyonu
def connect_to_can():
    result = pcan_basic.CAN_Initialize(PCAN_USBBUS1, PCAN_BAUD_250K)
    if result == 0:
        print("CAN hattına başarıyla bağlanıldı!")
        return True
    else:
        print(f"CAN hattına bağlanma başarısız. Hata kodu: {result}")
        return False

# Bağlantıyı kapatma fonksiyonu
def close_can_connection():
    pcan_basic.CAN_Uninitialize(PCAN_USBBUS1)
    print("CAN hattı kapatıldı.")

# CAN mesajı alma thread'i
class CANReceiverThread(QThread):
    message_received = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.running = False

    def run(self):
        self.running = True
        while self.running:
            can_msg = TPCANMsg()
            result = pcan_basic.CAN_Read(PCAN_USBBUS1, byref(can_msg), None)
            if result == 0:
                data_str = f"ID: {hex(can_msg.ID)}, Data: {[hex(byte) for byte in can_msg.DATA[:can_msg.LEN]]}"
                self.message_received.emit(data_str)
            QThread.msleep(100)  # Kısa bir bekleme

    def stop(self):
        self.running = False
        self.wait()

# EBS Hata Kodlarını Okuma Fonksiyonu (UDS ile)
def read_ebs_error_codes():
    try:
        # UDS için gerekli servisler
        uds_service_diagnostic_session = 0x10  # DiagnosticSessionControl servisi
        sub_function_default_session = 0x01  # Default diagnostic session
        uds_service_read_dtc = 0x19  # ReadDTCInformation servisi
        sub_function_active_dtc = 0x02  # Aktif DTC'leri okuma sub fonksiyonu
        
        # 1. Oturum başlatma (Default session)
        can_msg_session = TPCANMsg()
        can_msg_session.ID = 0x18DA0B55  # EBS sisteminin CAN ID'si
        can_msg_session.MSGTYPE = 0x02  # Standart mesaj tipi
        can_msg_session.LEN = 8
        can_msg_session.DATA[0] = uds_service_diagnostic_session
        can_msg_session.DATA[1] = sub_function_default_session
        for i in range(2, 8):
            can_msg_session.DATA[i] = 0x00  # Geri kalan baytlar sıfır

        # Oturum başlatma mesajı gönderme
        result_session = pcan_basic.CAN_Write(PCAN_USBBUS1, byref(can_msg_session))
        if result_session != 0:
            QMessageBox.critical(None, "Hata", f"Oturum başlatılamadı. Hata kodu: {result_session}")
            return

        print("UDS oturumu başlatıldı!")

        QThread.msleep(200)  # Kısa bir bekleme süresi (oturumun tam olarak başlaması için)

        # 2. EBS'den aktif hata kodlarını okuma isteği gönderme
        can_msg_dtc = TPCANMsg()
        can_msg_dtc.ID = 0x18DA0B55  # EBS sisteminin CAN ID'si
        can_msg_dtc.MSGTYPE = 0x02  # Standart mesaj tipi
        can_msg_dtc.LEN = 8
        can_msg_dtc.DATA[0] = uds_service_read_dtc
        can_msg_dtc.DATA[1] = sub_function_active_dtc
        for i in range(2, 8):
            can_msg_dtc.DATA[i] = 0x00  # Geri kalan baytlar sıfır

        # DTC okuma mesajı gönderme
        result_dtc = pcan_basic.CAN_Write(PCAN_USBBUS1, byref(can_msg_dtc))
        if result_dtc == 0:
            print("EBS Hata kodları okuma isteği gönderildi!")
        else:
            QMessageBox.critical(None, "Hata", f"Mesaj gönderme başarısız. Hata kodu: {result_dtc}")

        # EBS'den gelen yanıtı al ve ekrana bas
        all_dtc_codes=[]
        while True:
            QThread.msleep(200)  # Kısa bir bekleme süresi
            can_msg_response = TPCANMsg()
            result_response = pcan_basic.CAN_Read(PCAN_USBBUS1, byref(can_msg_response), None)
            if result_response == 0:
                    if can_msg_response.LEN==0:
                        break 
                    data_str = f"EBS Hata Yanıtı: ID: {hex(can_msg_response.ID)}, Data: {[hex(byte) for byte in can_msg_response.DATA[:can_msg_response.LEN]]}"
                    all_dtc_codes.append(data_str)
            else: 
                break
        if all_dtc_codes:
                return "\n".join(all_dtc_codes)
        else:                                                                      
            QMessageBox.critical(None, "Hata", f"EBS'den yanıt alınamadı. Hata kodu: {result_response}")

    except Exception as e:
        QMessageBox.critical(None, "Hata", f"Beklenmeyen hata: {str(e)}")

# GUI Sınıfı
class CANInterface(QDialog):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.receiver_thread = None
        self.received_data = []  # Alınan verileri saklamak için

        # CAN bağlantısını kur
        if not connect_to_can():
            QMessageBox.critical(self,"Hata","CAN bağlantısı kurulamadı!")
            sys.exit(1)

    def init_ui(self):
        self.setWindowTitle("CAN HABERLEŞMESİ")
        self.setGeometry(100, 100, 600, 500)

        # Minimize butonu ekle
        self.setWindowFlags(Qt.Window | Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)

        layout = QVBoxLayout()
        
        # Mesaj Gönderme Bölümü
        send_layout = QVBoxLayout()
        send_label = QLabel("Mesaj Gönderme")
        send_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        send_layout.addWidget(send_label)

        # Mesaj ID
        id_layout = QHBoxLayout()
        id_label = QLabel("Mesaj ID (Hex):")
        self.id_input = QLineEdit()
        self.id_input.setPlaceholderText("Örneğin: 100")
        id_layout.addWidget(id_label)
        id_layout.addWidget(self.id_input)
        send_layout.addLayout(id_layout)

        # Veri Uzunluğu
        len_layout = QHBoxLayout()
        len_label = QLabel("Veri Uzunluğu (0-8):")
        self.len_input = QLineEdit()
        self.len_input.setPlaceholderText("Örneğin: 4")
        len_layout.addWidget(len_label)
        len_layout.addWidget(self.len_input)
        send_layout.addLayout(len_layout)

        # Veri Byte'ları
        data_layout = QHBoxLayout()
        data_label = QLabel("Veri Byte'ları (Hex, Boşluk ile ayırın):")
        self.data_input = QLineEdit()
        self.data_input.setPlaceholderText("Örneğin: 1A 2B 3C 4D")
        data_layout.addWidget(data_label)
        data_layout.addWidget(self.data_input)
        send_layout.addLayout(data_layout)

        # Gönder Butonu
        self.send_button = QPushButton("Gönder")
        self.send_button.clicked.connect(self.send_can_message)
        send_layout.addWidget(self.send_button)

        # Mesaj Alma Bölümü
        receive_layout = QVBoxLayout()
        receive_label = QLabel("CAN Mesajları")
        receive_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        receive_layout.addWidget(receive_label)

        # Alınan Mesajlar
        self.received_messages = QTextEdit()
        self.received_messages.setReadOnly(True)
        receive_layout.addWidget(self.received_messages)

        # Mesaj Alma ve Durdurma Butonları
        buttons_layout = QHBoxLayout()
        self.read_button = QPushButton("Veri Alma Başlat")
        self.read_button.clicked.connect(self.start_receiving)
        self.stop_button = QPushButton("Veri Alma Durdur")
        self.stop_button.clicked.connect(self.stop_receiving)
        buttons_layout.addWidget(self.read_button)
        buttons_layout.addWidget(self.stop_button)
        receive_layout.addLayout(buttons_layout)

        # Verileri Kaydet Butonu
        self.save_button = QPushButton("Verileri Kaydet (Excel)")
        self.save_button.clicked.connect(self.save_data)
        receive_layout.addWidget(self.save_button)

        # CAN Bağlantısını Kapat Butonu
        self.close_button = QPushButton("Bağlantıyı Kapat")
        self.close_button.clicked.connect(self.close_can)
        receive_layout.addWidget(self.close_button)

        layout.addLayout(send_layout)
        layout.addLayout(receive_layout)

        self.receive_text = QTextEdit()
        self.receive_text.setReadOnly(True)
        receive_layout.addWidget(self.receive_text)
        
         # Hata Kodlarını Okuma Butonu
        self.read_dtc_button = QPushButton("EBS Hata Kodlarını Oku")
        self.read_dtc_button.clicked.connect(self.read_dtc_codes)
        layout.addWidget(self.read_dtc_button)

        self.setLayout(layout)

    # CAN Mesajı Gönderme Fonksiyonu
    def send_can_message(self):
        try:
            msg_id = int(self.id_input.text(), 16)
            msg_len = int(self.len_input.text())
            msg_data = [int(x, 16) for x in self.data_input.text().split()]

            can_msg = TPCANMsg()
            can_msg.ID = msg_id
            can_msg.MSGTYPE = 0x02  # Standart mesaj tipi
            can_msg.LEN = msg_len
            for i in range(msg_len):
                can_msg.DATA[i] = msg_data[i]

            result = pcan_basic.CAN_Write(PCAN_USBBUS1, byref(can_msg))
            if result == 0:
                QMessageBox.information(self, "Başarılı", "Mesaj başarıyla gönderildi!")
            else:
                QMessageBox.critical(self, "Hata", f"Mesaj gönderilemedi. Hata kodu: {result}")

        except ValueError:
            QMessageBox.critical(self, "Hata", "Geçersiz giriş! Lütfen verileri doğru formatta girin.")

    # CAN Bağlantısını Kapatma Fonksiyonu
    def close_can(self):
        if self.receiver_thread:
            self.receiver_thread.stop()
        close_can_connection()
    def start_receiving(self):
        if self.receiver_thread is None:
            self.receiver_thread = CANReceiverThread()
            self.receiver_thread.message_received.connect(self.handle_received_message)
            self.receiver_thread.start()

    def stop_receiving(self):
        if self.receiver_thread is not None:
            self.receiver_thread.stop()
            self.receiver_thread = None

    def handle_received_message(self, data):
        self.received_messages.append(data)
        self.received_data.append(data)  # Alınan verileri kaydet


    # EBS Hata Kodlarını Okuma Fonksiyonu
    def read_dtc_codes(self):
        dtc_response = read_ebs_error_codes()
        if dtc_response:
            self.receive_text.append(dtc_response)

    # Verileri Kaydetme Fonksiyonu (Excel formatında)
    def save_data(self):
        if not self.received_data:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek veri yok!")
            return

        filename, _ = QFileDialog.getSaveFileName(self, "Verileri Kaydet", "", "Excel Dosyaları (*.xlsx)")
        if filename:
            df = pd.DataFrame(self.received_data, columns=["CAN Mesajları"])
            df.to_excel(filename, index=False)
            QMessageBox.information(self, "Başarılı", "Veriler başarıyla kaydedildi!")

# Uygulamayı çalıştır
app = QApplication(sys.argv)
can_interface = CANInterface()
can_interface.show()
sys.exit(app.exec_())
