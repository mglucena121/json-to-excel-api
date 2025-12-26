from PySide6.QtCore import QObject, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QProgressBar,
)
import sys
from pathlib import Path

import requests
import pandas as pd


class Worker(QObject):
    finished = Signal()
    failed = Signal(str)
    progress = Signal(int, str)  # valor 0-100 e etapa

    def __init__(self, url: str, output_path: str) -> None:
        super().__init__()
        self.url = url
        self.output_path = output_path

    def run(self) -> None:
        try:
            self.progress.emit(5, "Conectando…")
            response = requests.get(self.url, timeout=60)
            response.raise_for_status()

            self.progress.emit(35, "Baixando dados…")
            data = response.json()

            self.progress.emit(65, "Processando JSON…")
            df = pd.json_normalize(data)

            self.progress.emit(85, "Gravando Excel…")
            df.to_excel(self.output_path, index=False)

            self.progress.emit(100, "Concluído")
            self.finished.emit()
        except Exception as exc:  # pragma: no cover - simples diálogo de erro
            self.failed.emit(str(exc))


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("JSON para Excel")

        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("URL da API que retorna JSON")

        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("Caminho de saída do arquivo .xlsx")

        self.browse_btn = QPushButton("Escolher…")
        self.generate_btn = QPushButton("Gerar Excel")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)

        url_layout = QVBoxLayout()
        url_layout.addWidget(QLabel("URL da API"))
        url_layout.addWidget(self.url_input)

        path_layout = QHBoxLayout()
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(self.browse_btn)

        main_layout = QVBoxLayout()
        main_layout.addLayout(url_layout)
        main_layout.addWidget(QLabel("Arquivo de saída (.xlsx)"))
        main_layout.addLayout(path_layout)
        main_layout.addWidget(self.generate_btn)
        main_layout.addWidget(self.progress_bar)

        self.setLayout(main_layout)

        self.browse_btn.clicked.connect(self.on_browse)
        self.generate_btn.clicked.connect(self.on_generate)

        # Valores padrão úteis
        self.url_input.setText(
            "http://maispetoficial.intelite.com.br:3002/api/Dependents/relatorioAtivas?access_token=jmwlJyJQ6n4aW8EbgpNUVnsQolkvIkePH1G8Wc7jYG95ZVyUB4NZbhheFXkyZlai"
        )
        self.path_input.setText(str(Path.cwd() / "dados.xlsx"))

        self.thread: QThread | None = None
        self.worker: Worker | None = None

    def on_browse(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar como",
            str(Path.cwd() / "dados.xlsx"),
            "Planilha do Excel (*.xlsx)",
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.path_input.setText(path)

    def on_generate(self) -> None:
        url = self.url_input.text().strip()
        output_path = self.path_input.text().strip()

        if not url:
            QMessageBox.warning(self, "Validação", "Informe a URL da API.")
            return
        if not output_path:
            QMessageBox.warning(self, "Validação", "Informe o caminho do arquivo de saída.")
            return

        self.generate_btn.setEnabled(False)
        self.worker = Worker(url, output_path)
        self.thread = QThread()
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_success)
        self.worker.failed.connect(self.on_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.on_thread_done)
        self.thread.start()

    def on_success(self) -> None:
        QMessageBox.information(self, "Concluído", "Arquivo gerado com sucesso.")

    def on_failed(self, message: str) -> None:
        QMessageBox.critical(self, "Erro", f"Falha ao gerar o arquivo:\n{message}")

    def on_thread_done(self) -> None:
        self.generate_btn.setEnabled(True)
        self.url_input.setEnabled(True)
        self.path_input.setEnabled(True)
        self.browse_btn.setEnabled(True)
        # Limpa worker/thread
        self.worker = None
        self.thread = None

    def on_progress(self, value: int, _stage: str) -> None:
        self.progress_bar.setValue(value)


def main() -> None:
    app = QApplication(sys.argv)
    # Estilo escuro moderno via stylesheet
    app.setStyleSheet(
        """
        QWidget { background-color: #121212; color: #E0E0E0; font-size: 14px; }
        QLineEdit { background: #1E1E1E; border: 1px solid #2E2E2E; border-radius: 6px; padding: 6px; }
        QLineEdit:focus { border: 1px solid #3E8EED; }
        QPushButton { background: #2B2B2B; border: 1px solid #3A3A3A; border-radius: 6px; padding: 8px 12px; }
        QPushButton:hover { background: #333333; }
        QPushButton:disabled { color: #777; border-color: #2A2A2A; }
        QLabel { color: #CFCFCF; }
        QProgressBar { background: #1E1E1E; border: 1px solid #2E2E2E; border-radius: 6px; text-align: center; }
        QProgressBar::chunk { background-color: #3E8EED; border-radius: 6px; }
        """
    )
    win = MainWindow()
    win.resize(600, 200)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()


