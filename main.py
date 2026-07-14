# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import re
import numbers
from pathlib import Path
from datetime import datetime

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
            QWidget,
)


DISPLAY_COLUMNS = [
    "순번", "행사코드", "출발일", "도착일", "회원코드", "회원명", "객실종류",
    "객실호수", "항공좌석", "예약상태", "이름", "영문성", "영문이름", "연락처",
    "생년월일", "주민번호 뒷자리", "성별", "구분", "상품가", "항만세", "환율추가금",
    "유류할증료", "선내팁", "객실추가", "항공추가", "비자(금액)", "비자(체크여부)",
    "기타항목", "추가금액", "할인금액", "추가/할인 총 금액", "총 판매금액", "메모",
    "여행자 추가수집정보", "여권번호", "여권발급일", "여권만료일", "등록일",
]

EXCEL_NUMERIC_COLUMNS = {
    "상품가", "항만세", "환율추가금", "유류할증료", "선내팁", "객실추가", "항공추가",
    "비자(금액)", "추가금액", "할인금액", "추가/할인 총 금액", "총 판매금액",
}
EXTRA_COLUMNS = ["회원번호 1", "회원번호 2", "회원번호 3", "회원번호 4", "여행상태"]


class TravelMemberWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("여행고객 CRM 등록 (항차별)")
        self.resize(1700, 950)

        self.setStyleSheet(
            """
            QWidget { font-family: 'D2Coding', 'Malgun Gothic'; font-size: 10pt; color: #172033; }
            QFrame#panel { background: #e8edf4; border: 1px solid #aeb9c9; border-radius: 8px; }
            QLabel#blockTitle { font-size: 14pt; font-weight: 700; color: #111827; padding: 2px 0 8px 0; }
            QLabel#fieldTitle { font-weight: 700; color: #1f2937; }
            QPushButton { min-height: 32px; padding: 0 14px; background: #294b73; color: white; border: 1px solid #1e3a5f; border-radius: 4px; font-weight: 700; }
            QPushButton:hover { background: #386a9f; }
            QLineEdit, QComboBox { min-height: 30px; background: #ffffff; color: #111827; border: 1px solid #8795aa; border-radius: 3px; padding: 2px 6px; }
            QLineEdit:read-only { background: #f8fafc; }
            QTableWidget { background: #ffffff; alternate-background-color: #eef3f8; color: #111827; gridline-color: #aeb9c9; border: 1px solid #8795aa; }
            QTableWidget::item:hover { background: #dbeafe; color: #111827; }
            QTableWidget::item:selected { background: #ffffff; color: #111827; border: 1px solid #3975b8; }
            QTableWidget::item:selected:active { background: #ffffff; color: #111827; }
            QHeaderView::section { background: #cbd7e6; color: #172033; padding: 6px; border: 1px solid #aeb9c9; font-weight: 700; }
            QProgressBar { min-height: 18px; background: #ffffff; color: #172033; border: 1px solid #8795aa; border-radius: 4px; text-align: center; }
            QProgressBar::chunk { background: #3975b8; border-radius: 3px; }
            QTextEdit { background: #101a10; color: #69e06f; border: 1px solid #253c25; }
            """
        )

        root = QWidget()
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(12)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._build_trip_panel())
        splitter.addWidget(self._build_detail_panel())
        splitter.setSizes([430, 1200])
        root_layout.addWidget(splitter)
        self.setCentralWidget(root)

        self._write_log("화면이 준비되었습니다.")
        self._write_log("좌측 항차 목록에서 행사코드를 선택하면 고객 자료가 표시됩니다.")

    def _build_trip_panel(self) -> QFrame:
        panel = QFrame(objectName="panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)

        title = QLabel("여행항차", objectName="blockTitle")
        layout.addWidget(title)

        self.trip_table = QTableWidget(0, 5)
        self.trip_table.setHorizontalHeaderLabels(["", "여행지역", "출발일", "도착일", "행사코드"])
        self.trip_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.trip_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.trip_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.trip_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.trip_table.verticalHeader().setVisible(False)
        self.trip_table.setColumnWidth(0, 34)
        self.trip_table.setColumnWidth(1, 120)
        self.trip_table.setColumnWidth(2, 100)
        self.trip_table.setColumnWidth(3, 100)
        self.trip_table.setColumnWidth(4, 150)
        self.trip_table.cellClicked.connect(self._select_trip)
        layout.addWidget(self.trip_table)
        return panel

    def _build_detail_panel(self) -> QFrame:
        panel = QFrame(objectName="panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        file_row = QHBoxLayout()
        file_row.addWidget(QLabel("파일검색", objectName="fieldTitle"))
        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("불러온 엑셀 파일 경로 및 파일명")
        file_row.addWidget(self.file_path, 1)
        file_button = QPushButton("① 파일검색")
        file_button.clicked.connect(self._browse_file)
        file_row.addWidget(file_button)
        query_button = QPushButton("② 조회")
        query_button.clicked.connect(self._query_file)
        file_row.addWidget(query_button)
        review_button = QPushButton("③ 검토")
        review_button.clicked.connect(self._review_data)
        apply_button = QPushButton("④ 적용")
        apply_button.clicked.connect(self._apply_data)
        download_button = QPushButton("📊 엑셀 다운로드")
        download_button.clicked.connect(self._export_to_excel)
        apply_button.setMinimumWidth(int(apply_button.sizeHint().width() * 1.5))
        numbered_button_width = apply_button.minimumWidth()
        for button in (file_button, query_button, review_button, apply_button):
            button.setFixedWidth(numbered_button_width)
        download_button.setMinimumWidth(int(download_button.sizeHint().width() * 1.5))
        layout.addLayout(file_row)

        event_row = QHBoxLayout()
        event_row.addWidget(QLabel("행사코드", objectName="fieldTitle"))
        self.event_code = QLineEdit()
        self.event_code.setPlaceholderText("조회된 행사코드 또는 수정할 행사코드")
        event_row.addWidget(self.event_code, 1)
        event_row.addWidget(QLabel("인원당 계정수", objectName="fieldTitle"))
        self.account_count = QComboBox()
        self.account_count.addItems(["", "1계정", "2계정", "3계정", "4계정"])
        self.account_count.setCurrentIndex(0)
        event_row.addWidget(self.account_count)
        event_row.addWidget(review_button)
        event_row.addWidget(apply_button)
        event_row.addWidget(download_button)
        layout.addLayout(event_row)

        self.member_table = QTableWidget(0, len(DISPLAY_COLUMNS))
        self.member_table.setHorizontalHeaderLabels(DISPLAY_COLUMNS)
        self.member_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.member_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.member_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.member_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.member_table.verticalHeader().setVisible(False)
        self.member_table.setAlternatingRowColors(True)
        layout.addWidget(self.member_table, 1)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setFormat("대기 중")
        layout.addWidget(self.progress)

        log_title = QLabel("처리현황", objectName="fieldTitle")
        layout.addWidget(log_title)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFont(QFont("Consolas", 10))
        self.log.setMinimumHeight(135)
        layout.addWidget(self.log)
        return panel

    def _add_trip_row(self, region: str, departure: str, arrival: str, event_code: str, processed: bool) -> None:
        row = self.trip_table.rowCount()
        self.trip_table.insertRow(row)
        status = QTableWidgetItem("●" if processed else "X")
        status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        status.setForeground(QColor("#2563eb" if processed else "#dc2626"))
        status.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        self.trip_table.setItem(row, 0, status)
        for col, value in enumerate((region, departure, arrival, event_code), start=1):
            self.trip_table.setItem(row, col, QTableWidgetItem(value))

    def _select_trip(self, row: int, _column: int) -> None:
        event_code = self.trip_table.item(row, 4).text()
        self.event_code.setText(event_code)
        self._load_preview_rows(event_code)
        self._write_log(f"항차 선택: {event_code}")

    def _load_preview_rows(self, event_code: str) -> None:
        self.member_table.setRowCount(0)
        preview = [
            ["1", event_code, "2026-07-14", "2026-07-24", "2405140052", "김수덕", "내측2인실", "10040", "", "예약 완료", "김수덕", "KIM", "SOODUCK", "010-2713-2510", "1958-10-24", "", "F", "성인", "5445000", "300000", "350000", "0", "0", "0", "0", "0", "", "0", "0", "0", "650000", "6095000", "", "", "M286X8923", "2023-09-15", "2033-09-15", "2026-03-20"],
        ]
        for values in preview:
            row = self.member_table.rowCount()
            self.member_table.insertRow(row)
            for col, value in enumerate(values):
                self.member_table.setItem(row, col, QTableWidgetItem(value))

    def _browse_file(self) -> None:
        executable_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
        initial_dir = executable_dir / "xlsx"
        if not initial_dir.is_dir():
            initial_dir = executable_dir
        path, _ = QFileDialog.getOpenFileName(
            self,
            "엑셀 파일 선택",
            str(initial_dir),
            "Excel Files (*.xlsx *.xls);;XLSX Files (*.xlsx);;XLS Files (*.xls)",
        )
        if path:
            self.file_path.setText(path)
            self._write_log(f"파일 선택: {path}")

    def _query_file(self) -> None:
        if not self.file_path.text().strip():
            QMessageBox.information(self, "조회", "먼저 엑셀 파일을 선택해줘.")
            return
        file_path = Path(self.file_path.text().strip())
        if not file_path.is_file():
            QMessageBox.warning(self, "조회", "선택한 파일을 찾을 수 없어.")
            self._write_log(f"오류: 파일이 없음 - {file_path}")
            return

        try:
            self.progress.setRange(0, 0)
            self.progress.setFormat("엑셀 조회 중...")
            self._load_excel(file_path)
            self.progress.setRange(0, 100)
            self.progress.setValue(100)
            self.progress.setFormat("조회 완료")
        except Exception as exc:
            self.progress.setRange(0, 100)
            self.progress.setValue(0)
            self.progress.setFormat("조회 실패")
            QMessageBox.critical(self, "조회 오류", f"엑셀 파일을 읽지 못했어.\n{exc}")
            self._write_log(f"오류: {exc}")

    def _load_excel(self, file_path: Path) -> None:
        engine = "xlrd" if file_path.suffix.lower() == ".xls" else "openpyxl"
        raw = pd.read_excel(file_path, header=None, engine=engine)
        raw = raw.dropna(how="all").reset_index(drop=True)
        if raw.empty:
            raise ValueError("엑셀에 데이터가 없어.")

        header_row = self._find_header_row(raw)
        headers = [self._clean_header(value, index) for index, value in enumerate(raw.iloc[header_row].tolist())]
        data = raw.iloc[header_row + 1:].copy()
        data.columns = headers
        data = data.dropna(how="all").fillna("")

        event_text = " ".join(str(value) for value in raw.iloc[:header_row].fillna("").values.flatten())
        event_code, departure, arrival = self._extract_trip_info(event_text)
        self.event_code.setText(event_code)
        data = data.drop(columns=["회원번호"], errors="ignore")
        leader_count = self._normalize_tour_leader(data)
        for extra_column in EXTRA_COLUMNS:
            data[extra_column] = ""
        self._set_member_table(data)

        self._write_log(f"엑셀 조회 완료: {len(data)}건 / 행사코드: {event_code or '미확인'}")
        self._write_log(f"출발일: {departure or '미확인'} / 도착일: {arrival or '미확인'}")
        if leader_count:
            self._write_log(f"TOUR LEADER 변환: {leader_count}건")

    @staticmethod
    def _normalize_tour_leader(data: pd.DataFrame) -> int:
        if "예약상태" not in data.columns or "회원코드" not in data.columns:
            return 0
        mask = data["예약상태"].astype(str).str.strip().str.upper().eq("TOUR LEADER")
        count = int(mask.sum())
        if count:
            data.loc[mask, "회원코드"] = "TOUR LEADER"
            data.loc[mask, "예약상태"] = ""
        return count

    @staticmethod
    def _find_header_row(raw: pd.DataFrame) -> int:
        for index, row in raw.iterrows():
            values = {str(value).strip() for value in row.tolist() if str(value).strip()}
            if "회원코드" in values or "회원번호" in values:
                return int(index)
        raise ValueError("회원코드 또는 회원번호가 있는 헤더 행을 찾지 못했어.")

    @staticmethod
    def _clean_header(value: object, index: int) -> str:
        text = str(value).strip()
        return text if text and text.lower() != "nan" else f"컬럼{index + 1}"

    @staticmethod
    def _extract_trip_info(text: str) -> tuple[str, str, str]:
        event_patterns = [
            r"(?:행사\s*코드|행사코드|이벤트\s*코드)\s*[:：]\s*([^/|,]+)",
        ]
        date_patterns = {
            "departure": r"(?:출발\s*일|출발일|출발\s*일자)\s*[:：]\s*([^/|,]+)",
            "arrival": r"(?:도착\s*일|도착일|도착\s*일자|귀국일)\s*[:：]\s*([^/|,]+)",
        }
        event_code = TravelMemberWindow._first_match(text, event_patterns)
        departure = TravelMemberWindow._first_match(text, [date_patterns["departure"]])
        arrival = TravelMemberWindow._first_match(text, [date_patterns["arrival"]])
        return event_code, TravelMemberWindow._normalize_date(departure), TravelMemberWindow._normalize_date(arrival)

    @staticmethod
    def _first_match(text: str, patterns: list[str]) -> str:
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return ""

    @staticmethod
    def _normalize_date(value: str) -> str:
        if not value:
            return ""
        cleaned = value.strip()
        match = re.search(r"(20\d{2})\s*[년./-]\s*(\d{1,2})\s*[월./-]\s*(\d{1,2})", cleaned)
        if match:
            return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}-{int(match.group(3)):02d}"
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
            try:
                return datetime.strptime(cleaned[:10], fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
        return cleaned

    def _set_member_table(self, data: pd.DataFrame) -> None:
        self.member_table.clear()
        self.member_table.setColumnCount(len(data.columns))
        self.member_table.setHorizontalHeaderLabels([str(column) for column in data.columns])
        for column_index, column in enumerate(data.columns):
            if str(column) in EXTRA_COLUMNS:
                header_item = self.member_table.horizontalHeaderItem(column_index)
                header_item.setBackground(QColor("#f4d58d"))
                header_item.setForeground(QColor("#3b2f12"))
        self.member_table.setRowCount(len(data))
        for row_index, (_, row) in enumerate(data.iterrows()):
            for column_index, value in enumerate(row.tolist()):
                header = str(data.columns[column_index])
                display_value, is_numeric = self._format_cell(value, header)
                item = QTableWidgetItem(display_value)
                if is_numeric:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.member_table.setItem(row_index, column_index, item)
        self.member_table.resizeColumnsToContents()

    @staticmethod
    def _format_cell(value: object, header: str) -> tuple[str, bool]:
        if value is None or pd.isna(value):
            return "", False
        if isinstance(value, (datetime, pd.Timestamp)):
            return value.strftime("%Y-%m-%d"), False
        if header in EXCEL_NUMERIC_COLUMNS:
            if isinstance(value, numbers.Number) and not isinstance(value, bool):
                number = float(value)
                if number.is_integer():
                    return f"{int(number):,}", True
                return f"{number:,.2f}".rstrip("0").rstrip("."), True
            text = str(value).strip()
            if re.fullmatch(r"-?\d+(?:\.\d+)?", text):
                number = float(text)
                if number.is_integer():
                    return f"{int(number):,}", True
                return f"{number:,.2f}".rstrip("0").rstrip("."), True
        return str(value), False

    def _review_data(self) -> None:
        if self.member_table.rowCount() == 0:
            self._write_log("검토할 엑셀 데이터가 없어.")
            QMessageBox.information(self, "③ 검토", "먼저 엑셀 파일을 조회해줘.")
            return
        self._write_log(f"검토 대상: {self.member_table.rowCount()}건")
        self._write_log("회원번호 1~4와 여행상태를 확인해줘.")

    def _apply_data(self) -> None:
        if self.member_table.rowCount() == 0:
            self._write_log("적용할 데이터가 없어.")
            QMessageBox.information(self, "④ 적용", "먼저 엑셀 파일을 조회해줘.")
            return
        self._write_log("④ 적용 요청을 받았어. DB 등록 로직은 다음 단계에서 연결할 예정이야.")

    def _export_to_excel(self) -> None:
        if self.member_table.rowCount() == 0:
            self._write_log("다운로드할 데이터가 없어.")
            QMessageBox.information(self, "엑셀 다운로드", "먼저 엑셀 파일을 조회해줘.")
            return
        default_dir = Path(self.file_path.text()).parent if self.file_path.text() else Path.cwd()
        save_path, _ = QFileDialog.getSaveFileName(
            self, "엑셀 파일 저장", str(default_dir / "여행고객_검토결과.xlsx"), "Excel Files (*.xlsx)"
        )
        if not save_path:
            return
        headers = [self.member_table.horizontalHeaderItem(col).text() for col in range(self.member_table.columnCount())]
        rows = []
        for row in range(self.member_table.rowCount()):
            rows.append([
                self.member_table.item(row, col).text() if self.member_table.item(row, col) else ""
                for col in range(self.member_table.columnCount())
            ])
        pd.DataFrame(rows, columns=headers).to_excel(save_path, index=False)
        self._write_log(f"엑셀 다운로드 완료: {save_path}")

    def _write_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log.append(f"{timestamp} [INFO] {message}")


def main() -> int:
    app = QApplication(sys.argv)
    app.setFont(QFont("D2Coding", 10))
    window = TravelMemberWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
