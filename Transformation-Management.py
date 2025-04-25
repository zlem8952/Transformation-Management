import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QHBoxLayout,
    QComboBox, QProgressBar, QTextEdit, QFileDialog, QMessageBox,
    QListView, QTreeView, QAbstractItemView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
from PIL import Image
import pandas as pd
import fitz  # PyMuPDF
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import concurrent.futures

# LibreOffice(soffice) 실행파일 경로를 명시적으로 지정
SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
# SOFFICE_PATH = r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"

class ConvertWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)  # 실패 파일 리스트를 넘김

    def __init__(self, folders, src_format, target_format, max_workers=4):
        super().__init__()
        self.folders = folders
        self.src_format = src_format.lower()
        self.target_format = target_format.lower()
        self.supported = {
            'pdf': ['.pdf'],
            'png': ['.png', '.jpg', '.jpeg'],
            'excel': ['.xls', '.xlsx']
        }
        self.failed_files = []
        self.max_workers = max_workers

    def run(self):
        file_list = self.find_files()
        total = len(file_list)
        if total == 0:
            self.progress.emit(0, "해당 형식의 파일이 없습니다.")
            self.finished.emit([])
            return

        # 멀티스레딩으로 파일 변환
        with concurrent.futures.ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = []
            for idx, file_path in enumerate(file_list):
                futures.append(executor.submit(self.convert_file, file_path, idx, total))
            for future in concurrent.futures.as_completed(futures):
                percent, msg, failed = future.result()
                self.progress.emit(percent, msg)
                if failed:
                    self.failed_files.append(failed)

        self.finished.emit(self.failed_files)

    def find_files(self):
        file_list = []
        for folder in self.folders:
            for dirpath, _, filenames in os.walk(folder):
                for f in filenames:
                    ext = os.path.splitext(f)[1].lower()
                    if ext in self.supported[self.src_format]:
                        file_list.append(os.path.join(dirpath, f))
        return file_list

    def get_unique_path(self, path):
        counter = 1
        new_path = path
        while os.path.exists(new_path):
            base, ext = os.path.splitext(path)
            new_path = f"{base}({counter}){ext}"
            counter += 1
        return new_path

    def convert_file(self, file_path, idx, total):
        base, ext = os.path.splitext(file_path)
        output_path = self.get_unique_path(f"{base}.{self.target_format}")
        failed = None
        try:
            # PDF 변환
            if self.src_format == 'pdf':
                if self.target_format == 'png':
                    doc = fitz.open(file_path)
                    for page_num, page in enumerate(doc):
                        pix = page.get_pixmap(dpi=300)
                        page_path = self.get_unique_path(f"{base}_page{page_num+1}.png")
                        pix.save(page_path)
                else:
                    raise NotImplementedError("PDF→Excel 변환은 지원하지 않습니다.")

            # 이미지 변환
            elif self.src_format == 'png':
                if self.target_format == 'pdf':
                    img = Image.open(file_path)
                    img = img.convert("RGB")
                    img.save(output_path, "PDF", resolution=300)
                else:
                    raise NotImplementedError("PNG→Excel 변환은 지원하지 않습니다.")

            # 엑셀 변환 (LibreOffice 사용, 경로 명시)
            elif self.src_format == 'excel':
                if self.target_format == 'pdf':
                    import subprocess
                    try:
                        cmd = [
                            SOFFICE_PATH,
                            '--headless',
                            '--convert-to', 'pdf',
                            '--outdir', os.path.dirname(file_path),
                            file_path
                        ]
                        subprocess.run(cmd, check=True)
                    except Exception as e:
                        raise Exception(f"LibreOffice 변환 실패: {e}")

            msg = f"[성공] {os.path.basename(file_path)}"
        except Exception as e:
            msg = f"[실패] {os.path.basename(file_path)}: {str(e)}"
            failed = (file_path, str(e))
        percent = int((idx + 1) / total * 100)
        return percent, msg, failed

class FileConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("파일 변환 마스터")
        self.setGeometry(300, 300, 800, 600)
        self.setWindowIcon(QIcon())

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        self.btn_folder = QPushButton("폴더 여러 개 선택")
        self.btn_folder.clicked.connect(self.select_folders)
        layout.addWidget(self.btn_folder)

        format_layout = QHBoxLayout()
        self.src_combo = QComboBox()
        self.src_combo.addItems(["PDF", "PNG", "Excel"])
        self.target_combo = QComboBox()
        self.update_target_combo()
        self.src_combo.currentIndexChanged.connect(self.update_target_combo)
        format_layout.addWidget(QLabel("원본 형식:"))
        format_layout.addWidget(self.src_combo)
        format_layout.addWidget(QLabel("변환 형식:"))
        format_layout.addWidget(self.target_combo)
        layout.addLayout(format_layout)

        self.btn_convert = QPushButton("변환 시작")
        self.btn_convert.setEnabled(False)
        self.btn_convert.clicked.connect(self.start_conversion)
        layout.addWidget(self.btn_convert)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        central_widget.setLayout(layout)
        self.selected_folders = []

    def select_folders(self):
        dialog = QFileDialog(self, "폴더 여러 개 선택")
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.ShowDirsOnly, True)
        dialog.setOption(QFileDialog.DontUseNativeDialog, True)
        # 다중 선택 활성화
        for view in dialog.findChildren(QListView) + dialog.findChildren(QTreeView):
            view.setSelectionMode(QAbstractItemView.MultiSelection)
        if dialog.exec_():
            folders = dialog.selectedFiles()
            if folders:
                self.selected_folders = folders
                self.btn_convert.setEnabled(True)
                self.log.append(f"선택된 폴더: {', '.join(folders)}")

    def update_target_combo(self):
        src = self.src_combo.currentText()
        self.target_combo.clear()
        targets = {
            'PDF': ['PNG'],
            'PNG': ['PDF'],
            'Excel': ['PDF']
        }[src]
        self.target_combo.addItems(targets)

    def start_conversion(self):
        self.btn_folder.setEnabled(False)
        self.btn_convert.setEnabled(False)
        self.progress.setValue(0)
        self.log.clear()

        self.worker = ConvertWorker(
            self.selected_folders,
            self.src_combo.currentText(),
            self.target_combo.currentText(),
            max_workers=4  # 동시 변환 워커 수 (CPU 코어 수에 맞게 조절)
        )
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def update_progress(self, percent, msg):
        self.progress.setValue(percent)
        self.log.append(msg)

    def on_finished(self, failed_files):
        self.btn_folder.setEnabled(True)
        self.btn_convert.setEnabled(True)
        if failed_files:
            msg = "\n".join([f"{os.path.basename(f)}: {err}" for f, err in failed_files])
            QMessageBox.warning(self, "변환 실패 파일", f"다음 파일 변환 실패:\n{msg}")
        else:
            QMessageBox.information(self, "완료", "모든 변환이 완료되었습니다!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = FileConverter()
    window.show()
    sys.exit(app.exec_())
