# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import re
import numbers
from pathlib import Path
from datetime import datetime

import pandas as pd
from DB_conn import connect
from PySide6.QtCore import QTimer, Qt
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
    "순번", "행사코드", "출발일", "도착일", "여행지역", "회원코드", "회원명", "객실종류",
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

DB_COLUMN_MAP = {
    "행사코드": "EventCode",
    "출발일": "DepartureDate",
    "도착일": "ArrivalDate",
    "여행지역": "TravelRegion",
    "회원코드": "MemberCode",
    "회원명": "MemberName",
    "객실종류": "RoomType",
    "객실호수": "RoomNo",
    "항공좌석": "AirlineSeat",
    "예약상태": "ReservationStatus",
    "이름": "CustomerName",
    "영문성": "EnglishLastName",
    "영문이름": "EnglishName",
    "연락처": "ContactNo",
    "생년월일": "BirthDate",
    "주민번호 뒷자리": "ResidentNoTail",
    "성별": "Gender",
    "구분": "CustomerType",
    "상품가": "ProductPrice",
    "항만세": "PortTax",
    "환율추가금": "ExchangeRateExtra",
    "유류할증료": "FuelSurcharge",
    "선내팁": "OnboardTip",
    "객실추가": "RoomExtra",
    "항공추가": "AirlineExtra",
    "비자(금액)": "VisaAmount",
    "비자(체크여부)": "VisaChecked",
    "기타항목": "OtherItem",
    "추가금액": "ExtraAmount",
    "할인금액": "DiscountAmount",
    "추가/할인 총 금액": "TotalExtraDiscount",
    "총 판매금액": "TotalSaleAmount",
    "메모": "Memo",
    "여행자 추가수집정보": "TravelerExtraInfo",
    "여권번호": "PassportNo",
    "여권발급일": "PassportIssueDate",
    "여권만료일": "PassportExpiryDate",
    "등록일": "RegisteredDate",
}

RIGHT_DB_COLUMNS = [
    "순번",
    *[
        column for column in DISPLAY_COLUMNS[1:]
        if column not in {"행사코드", "여행지역", "출발일", "도착일"}
    ],
]


class TravelMemberWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("여행고객 CRM 등록 (항차별)")
        self.resize(1700, 950)
        self._base_data_saved = False
        self._saved_parent_seq_by_row: list[int] = []

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
            QScrollBar:horizontal { height: 18px; background: #c8d1de; border: 1px solid #8795aa; }
            QScrollBar:vertical { width: 18px; background: #c8d1de; border: 1px solid #8795aa; }
            QScrollBar::handle:horizontal, QScrollBar::handle:vertical { background: #526b88; border: 1px solid #34495f; border-radius: 3px; min-width: 40px; min-height: 40px; }
            QScrollBar::handle:hover { background: #294b73; }
            QScrollBar::add-line, QScrollBar::sub-line { background: #9eacbd; border: 1px solid #718096; }
            QProgressBar { min-height: 18px; background: #ffffff; color: #172033; border: 1px solid #8795aa; border-radius: 4px; text-align: center; }
            QProgressBar::chunk { background: #3975b8; border-radius: 3px; }
            QTextEdit { background: #101a10; color: #69e06f; border: 1px solid #253c25; }
            QMessageBox { background: #f7f9fc; color: #111827; }
            QMessageBox QLabel { background: transparent; color: #111827; min-width: 300px; }
            QMessageBox QPushButton { background: #294b73; color: white; border: 1px solid #1e3a5f; border-radius: 4px; padding: 6px 18px; min-width: 80px; }
            QMessageBox QPushButton:hover { background: #386a9f; color: white; }
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
        QTimer.singleShot(0, self._load_trip_list_from_db)

    def _build_trip_panel(self) -> QFrame:
        panel = QFrame(objectName="panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)

        title = QLabel("여행항차", objectName="blockTitle")
        layout.addWidget(title)

        self.trip_table = QTableWidget(0, 6)
        self.trip_table.setHorizontalHeaderLabels(["순번", "", "여행지역", "출발일", "도착일", "행사코드"])
        self.trip_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.trip_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.trip_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.trip_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.trip_table.verticalHeader().setVisible(False)
        self.trip_table.setColumnWidth(0, 55)
        self.trip_table.setColumnWidth(1, 42)
        self.trip_table.setColumnWidth(2, 120)
        self.trip_table.setColumnWidth(3, 100)
        self.trip_table.setColumnWidth(4, 100)
        self.trip_table.setColumnWidth(5, 150)
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
        apply_button.setMinimumWidth(int(apply_button.sizeHint().width() * 1.2))
        numbered_button_width = apply_button.minimumWidth()
        for button in (file_button, query_button, review_button, apply_button):
            button.setFixedWidth(numbered_button_width)
        download_button.setMinimumWidth(int(download_button.sizeHint().width() * 1.2))
        layout.addLayout(file_row)

        event_row = QHBoxLayout()
        event_row.addWidget(QLabel("행사코드", objectName="fieldTitle"))
        self.event_code = QLineEdit()
        self.event_code.setPlaceholderText("조회된 행사코드 또는 수정할 행사코드")
        event_row.addWidget(self.event_code, 1)
        event_row.addWidget(QLabel("여행지역", objectName="fieldTitle"))
        self.travel_region = QLineEdit()
        self.travel_region.setReadOnly(True)
        self.travel_region.setPlaceholderText("여행지역")
        self.travel_region.setFixedWidth(120)
        event_row.addWidget(self.travel_region)
        event_row.addWidget(QLabel("출발일", objectName="fieldTitle"))
        self.departure_date = QLineEdit()
        self.departure_date.setReadOnly(True)
        self.departure_date.setPlaceholderText("YYYY-MM-DD")
        self.departure_date.setFixedWidth(112)
        event_row.addWidget(self.departure_date)
        event_row.addWidget(QLabel("도착일", objectName="fieldTitle"))
        self.arrival_date = QLineEdit()
        self.arrival_date.setReadOnly(True)
        self.arrival_date.setPlaceholderText("YYYY-MM-DD")
        self.arrival_date.setFixedWidth(112)
        event_row.addWidget(self.arrival_date)
        event_row.addWidget(QLabel("인원당 계정수", objectName="fieldTitle"))
        self.account_count = QComboBox()
        self.account_count.addItems(["", "1계정", "2계정", "3계정", "4계정"])
        self.account_count.setCurrentIndex(0)
        event_row.addWidget(self.account_count)
        event_row.addWidget(review_button)
        event_row.addWidget(apply_button)
        event_row.addWidget(download_button)
        layout.addLayout(event_row)

        self.member_table = QTableWidget(0, len(RIGHT_DB_COLUMNS))
        self.member_table.setHorizontalHeaderLabels(RIGHT_DB_COLUMNS)
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
        row_number = QTableWidgetItem(str(row + 1))
        row_number.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.trip_table.setItem(row, 0, row_number)
        status = QTableWidgetItem("●" if processed else "X")
        status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        status.setForeground(QColor("#2563eb" if processed else "#dc2626"))
        status.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        self.trip_table.setItem(row, 1, status)
        for col, value in enumerate((region, departure, arrival, event_code), start=1):
            self.trip_table.setItem(row, col + 1, QTableWidgetItem(value))

    def _select_trip(self, row: int, _column: int) -> None:
        event_code = self.trip_table.item(row, 5).text()
        departure = self.trip_table.item(row, 3).text()
        arrival = self.trip_table.item(row, 4).text()
        travel_region = self.trip_table.item(row, 2).text()
        self.event_code.setText(event_code)
        self.travel_region.setText("" if travel_region == "미등록" else travel_region)
        self.departure_date.setText(departure)
        self.arrival_date.setText(arrival)
        self._load_trip_members(event_code, departure, arrival)
        self._write_log(f"항차 선택: {event_code}")

    def _load_trip_list_from_db(self) -> None:
        try:
            self._write_log("DB에서 여행항차 목록을 조회해.")
            with connect() as connection:
                cursor = connection.cursor()
                cursor.execute(
                    """
                    SELECT m.EventCode, m.DepartureDate, m.ArrivalDate, m.TravelRegion,
                           COUNT(a.TourMemberAccountSeqNo) AS AccountCount
                    FROM dbo.Tour_Member_M AS m
                    LEFT JOIN dbo.Tour_Member_Account AS a
                      ON a.TourMemberSeqNo = m.SeqNo
                    GROUP BY m.EventCode, m.DepartureDate, m.ArrivalDate, m.TravelRegion
                    ORDER BY MAX(m.SeqNo) DESC
                    """
                )
                trips = cursor.fetchall()
            self.trip_table.setRowCount(0)
            for event_code, departure, arrival, travel_region, account_count in trips:
                self._add_trip_row(
                    str(travel_region or "미등록"),
                    self._format_date_value(departure),
                    self._format_date_value(arrival),
                    str(event_code or ""),
                    int(account_count or 0) > 0,
                )
            self._write_log(f"여행항차 목록 조회 완료: {len(trips)}건")
        except Exception as exc:
            self._write_log(f"여행항차 목록 조회 오류: {exc}")

    def _load_trip_members(self, event_code: str, departure: str, arrival: str) -> None:
        db_columns = ["SeqNo", *DB_COLUMN_MAP.values()]
        try:
            with connect() as connection:
                cursor = connection.cursor()
                cursor.execute(
                    f"""
                    SELECT {', '.join(f'm.{column}' for column in db_columns)}
                    FROM dbo.Tour_Member_M AS m
                    WHERE m.EventCode = ?
                      AND m.DepartureDate = ?
                      AND m.ArrivalDate = ?
                    ORDER BY m.SeqNo
                    """,
                    (event_code, departure, arrival),
                )
                member_rows = cursor.fetchall()
                cursor.execute(
                    """
                    SELECT TourMemberSeqNo, AccountID, AllocationOrder
                    FROM dbo.Tour_Member_Account
                    WHERE EventCode = ?
                      AND DepartureDate = ?
                      AND ArrivalDate = ?
                    ORDER BY TourMemberSeqNo, AllocationOrder
                    """,
                    (event_code, departure, arrival),
                )
                account_rows = cursor.fetchall()

            accounts: dict[int, dict[int, str]] = {}
            for parent_seq, account_id, allocation_order in account_rows:
                accounts.setdefault(int(parent_seq), {})[int(allocation_order)] = str(account_id)

            headers = [*RIGHT_DB_COLUMNS, *EXTRA_COLUMNS]
            self.member_table.clear()
            self.member_table.setColumnCount(len(headers))
            self.member_table.setHorizontalHeaderLabels(headers)
            for column_index, header in enumerate(headers):
                if header in EXTRA_COLUMNS:
                    header_item = self.member_table.horizontalHeaderItem(column_index)
                    header_item.setBackground(QColor("#b7e4c7"))
                    header_item.setForeground(QColor("#123524"))
            self.member_table.setRowCount(len(member_rows))

            for row_index, values in enumerate(member_rows):
                parent_seq = int(values[0])
                row_values = list(values[1:])
                row_dict = dict(zip(DB_COLUMN_MAP.keys(), row_values))
                display_values = [
                    row_index + 1,
                    *[row_dict.get(column, "") for column in RIGHT_DB_COLUMNS[1:]],
                ]
                allocation = accounts.get(parent_seq, {})
                display_values.extend([allocation.get(order, "") for order in range(1, 5)])
                display_values.append("노쇼" if "노쇼" in str(row_dict.get("메모", "") or "") else "")
                for column_index, value in enumerate(display_values):
                    header = headers[column_index]
                    if header == "주민번호 뒷자리":
                        value = ""
                    display_value, is_numeric = self._format_cell(value, header)
                    item = QTableWidgetItem(display_value)
                    if is_numeric:
                        item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    self.member_table.setItem(row_index, column_index, item)
            self.member_table.resizeColumnsToContents()
            self._write_log(f"고객 자료 조회 완료: {len(member_rows)}건")
        except Exception as exc:
            QMessageBox.critical(self, "고객 자료 조회 오류", str(exc))
            self._write_log(f"고객 자료 조회 오류: {exc}")

    @staticmethod
    def _format_date_value(value: object) -> str:
        if value is None:
            return ""
        if isinstance(value, (datetime, pd.Timestamp)):
            return value.strftime("%Y-%m-%d")
        return str(value)[:10]

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
            self.progress.setStyleSheet(
                "QProgressBar { min-height: 18px; background: #3975b8; color: white; "
                "border: 1px solid #294b73; border-radius: 4px; text-align: center; font-weight: 700; } "
                "QProgressBar::chunk { background: #3975b8; border-radius: 3px; }"
            )
        except Exception as exc:
            self.progress.setRange(0, 100)
            self.progress.setValue(0)
            self.progress.setFormat("조회 실패")
            self.progress.setStyleSheet("")
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
        event_code, departure, arrival, travel_region = self._extract_trip_info(event_text)
        self.event_code.setText(event_code)
        self.travel_region.setText(travel_region)
        self.departure_date.setText(departure)
        self.arrival_date.setText(arrival)
        data = data.drop(columns=["회원번호"], errors="ignore")
        if "여행지역" not in data.columns:
            data.insert(0, "여행지역", travel_region)
        elif travel_region:
            data["여행지역"] = data["여행지역"].replace("", travel_region).fillna(travel_region)
        leader_count = self._normalize_tour_leader(data)
        for extra_column in EXTRA_COLUMNS:
            data[extra_column] = ""
        no_show_count = self._normalize_no_show(data)
        self._set_member_table(data)

        self._write_log(f"엑셀 조회 완료: {len(data)}건 / 행사코드: {event_code or '미확인'}")
        self._write_log(f"출발일: {departure or '미확인'} / 도착일: {arrival or '미확인'}")
        if leader_count:
            self._write_log(f"TOUR LEADER 변환: {leader_count}건")
        if no_show_count:
            self._write_log(f"노쇼 여행상태 반영: {no_show_count}건")
        self._base_data_saved = False
        self._saved_parent_seq_by_row = []

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
    def _normalize_no_show(data: pd.DataFrame) -> int:
        memo_column = next((column for column in ("메모", "Memo") if column in data.columns), None)
        if memo_column is None or "여행상태" not in data.columns:
            return 0
        mask = data[memo_column].astype(str).str.contains("노쇼", na=False)
        count = int(mask.sum())
        if count:
            data.loc[mask, "여행상태"] = "노쇼"
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
    def _extract_trip_info(text: str) -> tuple[str, str, str, str]:
        event_patterns = [
            r"(?:행사\s*코드|행사코드|이벤트\s*코드)\s*[:：]\s*([^/|,]+)",
        ]
        date_patterns = {
            "departure": r"(?:출발\s*일|출발일|출발\s*일자)\s*[:：]\s*([^/|,]+)",
            "arrival": r"(?:도착\s*일|도착일|도착\s*일자|귀국일)\s*[:：]\s*([^/|,]+)",
        }
        region = TravelMemberWindow._first_match(
            text,
            [r"(?:여행\s*지역|여행지역|여행지|지역)\s*[:：]\s*([^/|,]+)"],
        )
        event_code = TravelMemberWindow._first_match(text, event_patterns)
        departure = TravelMemberWindow._first_match(text, [date_patterns["departure"]])
        arrival = TravelMemberWindow._first_match(text, [date_patterns["arrival"]])
        return event_code, TravelMemberWindow._normalize_date(departure), TravelMemberWindow._normalize_date(arrival), region

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
        display_data = data.drop(columns=["여행지역"], errors="ignore")
        self.member_table.clear()
        headers = ["순번", *[str(column) for column in display_data.columns]]
        self.member_table.setColumnCount(len(headers))
        self.member_table.setHorizontalHeaderLabels(headers)
        for column_index, column in enumerate(display_data.columns):
            if str(column) in EXTRA_COLUMNS:
                header_item = self.member_table.horizontalHeaderItem(column_index + 1)
                header_item.setBackground(QColor("#b7e4c7"))
                header_item.setForeground(QColor("#123524"))
        self.member_table.setRowCount(len(display_data))
        for row_index in range(len(display_data)):
            row_number = QTableWidgetItem(str(row_index + 1))
            row_number.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.member_table.setItem(row_index, 0, row_number)
        for row_index, (_, row) in enumerate(display_data.iterrows()):
            for column_index, value in enumerate(row.tolist()):
                header = str(display_data.columns[column_index])
                display_value, is_numeric = self._format_cell(value, header)
                item = QTableWidgetItem(display_value)
                if is_numeric:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.member_table.setItem(row_index, column_index + 1, item)
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
        account_text = self.account_count.currentText().strip()
        if not account_text:
            QMessageBox.warning(self, "③ 검토", "인원당 계정수를 선택해야 검토를 진행할 수 있어.")
            self._write_log("검토 중단: 인원당 계정수가 선택되지 않았어.")
            return

        try:
            validation_errors = self._validate_base_data()
            if validation_errors:
                self._write_log(f"검토 중단: 기초자료 오류 {len(validation_errors)}건")
                QMessageBox.warning(
                    self,
                    "③ 검토 오류",
                    "조회된 엑셀 자료를 먼저 확인해줘.\n\n- " + "\n- ".join(validation_errors[:15]),
                )
                return

            if not self._base_data_saved:
                answer = QMessageBox.question(
                    self,
                    "기초자료 저장 확인",
                    "조회된 엑셀 자료에 이상이 없어.\n\n"
                    "기초자료를 데이터베이스에 먼저 저장 해야 합니다.\n\n"
                    "저장하고 검토를 시작할까요?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    self._write_log("검토 대기: 사용자가 기초자료 저장을 취소했어.")
                    return
                self._save_base_data()
                self._base_data_saved = True
                self._load_trip_list_from_db()

            account_count = int(account_text.replace("계정", ""))
            member_index = self._table_column_index("회원코드")
            code_values = {
                self.member_table.item(row, member_index).text().strip().lower()
                for row in range(self.member_table.rowCount())
                if self.member_table.item(row, member_index)
                and self.member_table.item(row, member_index).text().strip()
            }
            member_codes = sorted(code for code in code_values if code != "tour leader")
            self._write_log(f"검토 대상: {self.member_table.rowCount()}건 / 회원코드: {len(member_codes)}개")
            if not member_codes:
                self._write_log("검토 완료: DB 조회가 필요한 일반 회원코드가 없어.")
                return

            self.progress.setRange(0, 0)
            self.progress.setFormat("DB 검토 중...")
            member_dict, quantity_dict = self._load_member_accounts(member_codes)
            errors = self._allocate_accounts(member_dict, quantity_dict, account_count, member_index)
            self.progress.setRange(0, 100)
            self.progress.setValue(100)
            self.progress.setFormat("검토 완료")
            self.progress.setStyleSheet(
                "QProgressBar { min-height: 18px; background: #3975b8; color: white; "
                "border: 1px solid #294b73; border-radius: 4px; text-align: center; font-weight: 700; } "
                "QProgressBar::chunk { background: #3975b8; border-radius: 3px; }"
            )
            if errors:
                self._write_log(f"검토 완료: 오류 {errors}건")
            else:
                self._write_log("검토 완료: 계정 배분에 문제가 없어.")
        except Exception as exc:
            self.progress.setRange(0, 100)
            self.progress.setValue(0)
            self.progress.setFormat("검토 실패")
            self.progress.setStyleSheet("")
            QMessageBox.critical(self, "③ 검토 오류", str(exc))
            self._write_log(f"검토 오류: {exc}")

    def _validate_base_data(self) -> list[str]:
        errors: list[str] = []
        event_code = self.event_code.text().strip()
        departure = self.departure_date.text().strip()
        arrival = self.arrival_date.text().strip()
        if not event_code:
            errors.append("행사코드가 없어.")
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", departure):
            errors.append("출발일 형식이 YYYY-MM-DD가 아니야.")
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", arrival):
            errors.append("도착일 형식이 YYYY-MM-DD가 아니야.")

        try:
            member_index = self._table_column_index("회원코드")
            name_index = self._table_column_index("이름")
        except ValueError as exc:
            return errors + [str(exc)]

        for row in range(self.member_table.rowCount()):
            member_code = self._table_text(row, member_index).strip()
            customer_name = self._table_text(row, name_index).strip()
            if not member_code:
                errors.append(f"{row + 1}행: 회원코드가 없어.")
            if not customer_name and member_code.upper() != "TOUR LEADER":
                errors.append(f"{row + 1}행: 이름이 없어.")
        return errors

    def _save_base_data(self) -> None:
        target_columns = list(DB_COLUMN_MAP.values())
        insert_sql = f"""
            INSERT INTO dbo.Tour_Member_M
            ({', '.join(target_columns)})
            OUTPUT INSERTED.SeqNo
            VALUES ({', '.join('?' for _ in target_columns)})
        """
        headers = [
            self.member_table.horizontalHeaderItem(column).text()
            for column in range(self.member_table.columnCount())
        ]
        header_indexes = {header: index for index, header in enumerate(headers)}
        source_values = []
        for row in range(self.member_table.rowCount()):
            values = []
            for source_header, target_column in DB_COLUMN_MAP.items():
                if source_header == "행사코드":
                    value = self.event_code.text().strip()
                elif source_header == "출발일":
                    value = self.departure_date.text().strip()
                elif source_header == "도착일":
                    value = self.arrival_date.text().strip()
                elif source_header == "여행지역":
                    value = self.travel_region.text().strip()
                else:
                    value = self._table_text(row, header_indexes.get(source_header))
                values.append(self._db_value(source_header, value))
            source_values.append(values)

        self._write_log(f"Tour_Member_M 저장을 시작해. 대상 {len(source_values)}건")
        saved_seq_nos: list[int] = []
        with connect() as connection:
            cursor = connection.cursor()
            for values in source_values:
                cursor.execute(insert_sql, values)
                result = cursor.fetchone()
                if result is None:
                    raise RuntimeError("Tour_Member_M 저장 후 SeqNo를 확인하지 못했어.")
                saved_seq_nos.append(int(result[0]))
            connection.commit()
        self._saved_parent_seq_by_row = saved_seq_nos
        self._write_log(f"Tour_Member_M 저장 완료: {len(source_values)}건")

    def _db_value(self, header: str, value: str) -> str | int | None:
        if header == "주민번호 뒷자리":
            return None
        text = str(value).strip()
        if not text:
            return None
        if header in EXCEL_NUMERIC_COLUMNS:
            numeric_text = text.replace(",", "")
            if re.fullmatch(r"-?\d+(?:\.\d+)?", numeric_text):
                return float(numeric_text) if "." in numeric_text else int(numeric_text)
        return text

    def _table_text(self, row: int, column: int | None) -> str:
        if column is None:
            return ""
        item = self.member_table.item(row, column)
        return item.text() if item else ""

    def _load_member_accounts(self, member_codes: list[str]) -> tuple[dict[str, list[str]], dict[str, int]]:
        placeholders = ",".join("?" for _ in member_codes)
        query = f"""
            SELECT LOWER(me.memberno), LOWER(me.id)
            FROM member AS me
            INNER JOIN goods AS gu ON me.goods = gu.Goods_ID
            WHERE me.memtype IN (?, ?)
              AND LOWER(me.memberno) IN ({placeholders})
            UNION
            SELECT LOWER(me.memberno), LOWER(me.id)
            FROM member AS me
            LEFT JOIN goods AS gu ON me.Goods = gu.goods_id
            WHERE gu.Goods_ID = ?
              AND me.memtype <> ?
              AND LOWER(me.memberno) IN ({placeholders})
            ORDER BY 1, 2
        """
        params = ["정상", "만기", *member_codes, "undecided", "행사", *member_codes]
        quantity_query = f"""
            SELECT LOWER(me.memberno), 4 / NULLIF(CONVERT(INT, gu.G_etc_str5), 0)
            FROM member AS me
            INNER JOIN goods AS gu ON me.goods = gu.Goods_ID
            WHERE me.memtype IN (?, ?)
              AND LOWER(me.memberno) IN ({placeholders})
            UNION
            SELECT LOWER(me.memberno), 4 / NULLIF(CONVERT(INT, gu.G_etc_str5), 0)
            FROM member AS me
            LEFT JOIN goods AS gu ON me.Goods = gu.goods_id
            WHERE gu.Goods_ID = ?
              AND me.memtype <> ?
              AND LOWER(me.memberno) IN ({placeholders})
        """
        quantity_params = params
        self._write_log("DB 연결 및 회원 계정 조회를 시작해.")
        with connect() as connection:
            cursor = connection.cursor()
            cursor.execute(query, params)
            rows = cursor.fetchall()
            cursor.execute(quantity_query, quantity_params)
            quantity_rows = cursor.fetchall()

        member_dict: dict[str, list[str]] = {}
        for member_code, account_id in rows:
            member_dict.setdefault(str(member_code).lower(), []).append(str(account_id))
        quantity_dict: dict[str, int] = {}
        for member_code, quantity in quantity_rows:
            if quantity is not None:
                quantity_dict[str(member_code).lower()] = max(1, int(quantity))
        self._write_log(f"DB 조회 완료: 회원 {len(member_dict)}개 / 계정 {sum(len(v) for v in member_dict.values())}개")
        return member_dict, quantity_dict

    def _allocate_accounts(
        self,
        member_dict: dict[str, list[str]],
        quantity_dict: dict[str, int],
        account_count: int,
        member_index: int,
    ) -> int:
        account_columns = [self._table_column_index(name) for name in ("회원번호 1", "회원번호 2", "회원번호 3", "회원번호 4")]
        errors = 0
        for row in range(self.member_table.rowCount()):
            code_item = self.member_table.item(row, member_index)
            member_code = code_item.text().strip().lower() if code_item else ""
            if not member_code or member_code == "tour leader":
                continue
            if member_code not in member_dict or member_code not in quantity_dict:
                self._write_log(f"오류: 회원코드 {member_code}의 계정 정보를 찾지 못했어. (행 {row + 1})")
                errors += 1
                continue
            needed = account_count // quantity_dict[member_code]
            if needed <= 0 or len(member_dict[member_code]) < needed:
                self._write_log(f"오류: 계정 부족 - {member_code} (필요 {needed}개, 보유 {len(member_dict[member_code])}개)")
                errors += 1
                continue
            for column_index in account_columns:
                self.member_table.setItem(row, column_index, QTableWidgetItem(""))
            for offset in range(min(needed, len(account_columns))):
                self.member_table.setItem(row, account_columns[offset], QTableWidgetItem(member_dict[member_code].pop(0)))

        leftovers = [code for code, accounts in member_dict.items() if accounts]
        if leftovers:
            self._write_log(f"오류: 계정 분리가 필요한 회원코드 - {', '.join(leftovers)}")
            errors += len(leftovers)
        return errors

    def _table_column_index(self, header: str) -> int:
        for column in range(self.member_table.columnCount()):
            item = self.member_table.horizontalHeaderItem(column)
            if item and item.text() == header:
                return column
        raise ValueError(f"테이블에서 '{header}' 컬럼을 찾지 못했어.")

    def _apply_data(self) -> None:
        if self.member_table.rowCount() == 0:
            self._write_log("적용할 데이터가 없어.")
            QMessageBox.information(self, "④ 적용", "먼저 엑셀 파일을 조회해줘.")
            return
        if not self._base_data_saved or len(self._saved_parent_seq_by_row) != self.member_table.rowCount():
            QMessageBox.warning(self, "④ 적용", "먼저 ‘③ 검토’를 완료해서 기초자료를 저장해줘.")
            return
        account_columns = [self._table_column_index(name) for name in EXTRA_COLUMNS[:4]]
        if not any(self._table_text(row, column).strip() for row in range(self.member_table.rowCount()) for column in account_columns):
            QMessageBox.warning(self, "④ 적용", "먼저 ‘③ 검토’를 실행해서 회원번호를 할당해줘.")
            return
        answer = QMessageBox.question(
            self, "④ 적용 확인",
            "할당된 회원번호를 저장하고 행사 등록을 진행할까?\n\n"
            "Tour_Member_Account, Event, Event_Traveler, Event_Expenses 및 회원상태가 함께 반영돼.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            self.progress.setRange(0, 0)
            self.progress.setFormat("적용 중...")
            self._apply_database_transaction(account_columns)
            self.progress.setRange(0, 100)
            self.progress.setValue(100)
            self.progress.setFormat("적용 완료")
            self._load_trip_list_from_db()
            self._write_log("④ 적용 완료: 모든 대상 테이블에 저장했어.")
            QMessageBox.information(self, "④ 적용 완료", "행사 등록과 비용자료 저장이 완료됐어.")
        except Exception as exc:
            self.progress.setRange(0, 100)
            self.progress.setValue(0)
            self.progress.setFormat("적용 실패")
            self._write_log(f"④ 적용 실패(전체 롤백): {exc}")
            QMessageBox.critical(self, "④ 적용 오류", f"적용 중 오류가 발생해서 전체 취소했어.\n{exc}")

    def _apply_database_transaction(self, account_columns: list[int]) -> None:
        headers = [self.member_table.horizontalHeaderItem(column).text() for column in range(self.member_table.columnCount())]
        header_indexes = {header: index for index, header in enumerate(headers)}
        event_code = self.event_code.text().strip()
        departure = self.departure_date.text().strip()
        arrival = self.arrival_date.text().strip()
        region = self.travel_region.text().strip()
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with connect() as connection:
            cursor = connection.cursor()
            current_row = 0
            current_account = ""
            current_stage = "적용 준비"
            try:
                goods_cache: dict[tuple[str, str, str], str] = {}
                processed_ids: set[str] = set()
                applied_accounts = 0
                for row in range(self.member_table.rowCount()):
                    current_row = row + 1
                    parent_seq = self._saved_parent_seq_by_row[row]
                    customer_name = self._table_text(row, header_indexes.get("이름")) or self._table_text(row, header_indexes.get("회원명"))
                    member_code = self._table_text(row, header_indexes.get("회원코드"))
                    phone = self._table_text(row, header_indexes.get("연락처"))
                    travel_status = self._table_text(row, header_indexes.get("여행상태")) or "여행"
                    customer_values = {
                        "영문성": self._table_text(row, header_indexes.get("영문성")),
                        "영문이름": self._table_text(row, header_indexes.get("영문이름")),
                        "생년월일": self._table_text(row, header_indexes.get("생년월일")),
                        "객실종류": self._table_text(row, header_indexes.get("객실종류")),
                        "여권번호": self._table_text(row, header_indexes.get("여권번호")),
                        "여권만료일": self._table_text(row, header_indexes.get("여권만료일")),
                    }
                    expense_values = [
                        self._db_value("항만세", self._table_text(row, header_indexes.get("항만세"))),
                        self._db_value("유류할증료", self._table_text(row, header_indexes.get("유류할증료"))),
                        self._db_value("환율추가금", self._table_text(row, header_indexes.get("환율추가금"))),
                        self._db_value("선내팁", self._table_text(row, header_indexes.get("선내팁"))),
                        self._db_value("객실추가", self._table_text(row, header_indexes.get("객실추가"))),
                        self._db_value("항공추가", self._table_text(row, header_indexes.get("항공추가"))),
                        self._db_value("비자(금액)", self._table_text(row, header_indexes.get("비자(금액)"))),
                        self._db_value("추가금액", self._table_text(row, header_indexes.get("추가금액"))),
                    ]
                    for order, account_column in enumerate(account_columns, start=1):
                        account_id = self._table_text(row, account_column).strip()
                        if not account_id:
                            continue
                        current_account = account_id
                        current_stage = "회원번호 배정 저장"
                        cursor.execute(
                            """INSERT INTO dbo.Tour_Member_Account
                            (TourMemberSeqNo, EventCode, DepartureDate, ArrivalDate,
                             MemberCodeSnapshot, CustomerNameSnapshot, AccountID,
                             AllocationOrder, AllocationStatus, MatchMethod, AllocatedAt, AllocatedBy)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'ALLOCATED', 'MemberCode+Name', ?, 'Admin')""",
                            parent_seq, event_code, departure, arrival, member_code,
                            customer_name, account_id, order, now,
                        )
                        goods_key = (region, departure, arrival)
                        et_gid = goods_cache.get(goods_key)
                        if et_gid is None:
                            current_stage = "여행상품 저장"
                            et_gid = self._get_or_create_et_goods(cursor, event_code, region, departure, arrival, now)
                            goods_cache[goods_key] = et_gid
                        current_stage = "행사번호 확인"
                        result_kind, set_num = self._get_event_num(cursor, account_id, departure)
                        if result_kind == "NEW":
                            current_stage = "행사 저장"
                            self._insert_event(cursor, account_id, customer_name, departure, et_gid, set_num, event_code, travel_status, now)
                        current_stage = "여행자 저장"
                        self._insert_event_traveler(cursor, account_id, set_num, customer_name, phone, customer_values, now)
                        current_stage = "행사비용 저장"
                        self._upsert_event_expenses(cursor, account_id, set_num, expense_values, now)
                        if account_id not in processed_ids:
                            current_stage = "회원상태 및 메모 저장"
                            member_status = "노쇼" if travel_status == "노쇼" else "행사"
                            self._update_member_status(cursor, account_id, departure, member_status, now)
                            processed_ids.add(account_id)
                        applied_accounts += 1
                if applied_accounts == 0:
                    raise ValueError("저장할 회원번호가 없어.")
                connection.commit()
                self._write_log(f"Tour_Member_Account 및 행사자료 저장 완료: {applied_accounts}건")
            except Exception as exc:
                connection.rollback()
                raise RuntimeError(
                    f"{current_row}행 / 회원번호 {current_account or '-'} / {current_stage} 단계에서 실패했어.\n{exc}"
                ) from exc

    @staticmethod
    def _get_or_create_et_goods(cursor, event_code: str, region: str, departure: str, arrival: str, now: str) -> str:
        cursor.execute("SELECT TOP 1 ET_GID FROM dbo.ET_Goods WHERE SDate=? AND EDate=? AND ET_GName=?", departure, arrival, region)
        found = cursor.fetchone()
        if found:
            return str(found[0])
        cursor.execute("SELECT MAX(CAST(SUBSTRING(ET_GID, 5, 4) AS INT)) FROM dbo.ET_Goods WHERE LEN(ET_GID)=8")
        max_serial = cursor.fetchone()[0] or 0
        et_gid = f"{departure[:4]}{int(max_serial) + 1:04d}"
        columns = "ET_GID,ET_GName,ET_SDate,ET_EDate,ET_Type,ET_Kind,GMoney,GTMoney,ET_Cnt,ET_MinCnt,ET_MaxCnt,ET_GCnt,SDate,STime,SPos,EDate,ETime,EPos,ET_G,GI_SDate,GI_EDate,Traveler_IYN,ICompany,Tr_GuideYN,Local_GuideYN,Local_TrCompanyYN,SukPark,Hotel_Lev,Etc_SukPark,Local_Traffic,Bus_inseung,Etc_Traffic,meal_mark,Breakfast,Lunch,Dinner,Airplane_Seat,Ship_Seat,Train_Seat,TourCode,ET_JHGubun,Remark,CSuDang,Room_Upgrade,Tour_iGum,ET_FreeGum,etc_str1,etc_str2,etc_str3,etc_str4,etc_str5,Save_Date,Modi_Date,UsrID"
        values = [et_gid, region, departure, arrival, "0", "여행", 0, 0, 0, 0, 0, 0, departure, "", "", arrival, "", "", "", "", "", "", "", "", "", "", "", "", "", "", 0, "", "", 0, 0, 0, "", "", "", "", "진행", "", 0, 0, 0, 0, "", "", "", "", event_code, now, now, "Admin"]
        cursor.execute(f"INSERT INTO dbo.ET_Goods ({columns}) VALUES ({','.join('?' for _ in values)})", values)
        return et_gid

    @staticmethod
    def _get_event_num(cursor, account_id: str, departure: str) -> tuple[str, int]:
        cursor.execute("SELECT TOP 1 Num FROM dbo.Event WHERE ID=? ORDER BY Num DESC", account_id)
        found = cursor.fetchone()
        if found:
            return "OLD", int(found[0])
        prefix = f"{departure[2:4]}{departure[5:7]}"
        cursor.execute("SELECT ISNULL(MAX(CAST(RIGHT(CAST(Num AS varchar(20)),4) AS INT)),0) FROM dbo.Event WHERE LEN(CAST(Num AS varchar(20)))=8 AND LEFT(CAST(Num AS varchar(20)),4)=?", prefix)
        return "NEW", int(prefix) * 10000 + int(cursor.fetchone()[0] or 0) + 1

    @staticmethod
    def _insert_event(cursor, account_id: str, customer_name: str, departure: str, et_gid: str, set_num: int, event_code: str, travel_status: str, now: str) -> None:
        columns = "ID,Name,ECode,EName,EDate,ETime,EventType,ECharge_ID,EADate,EATime,EventPos,Remark,LunarEBDate,LeapMonth,EBDate,EBTimeGu,EBTime,Photo_Date,Photo_Time,EDCharge_ID,EDStep_ID1,EDStep_ID2,EDStep_ID3,Relation,Bigo,EReg_Date,SMS_EventYN,BalInDate,JangMooType,JangMoo,ShroudType,Shroud,ImJongJi,JangJi,Sex,SangJu,SangJuTel,Num,etc_str1,etc_str2,etc_str3,etc_str4,etc_str5,Area,Used_Money,cGoodsCnt,TravelPay,EGubun"
        values = [account_id, customer_name, "완료", customer_name, departure, "", "노쇼" if travel_status == "노쇼" else "여행", "", departure, "", et_gid, "", "", "0", "", "", "", "", "", "여행", "여행", "회원", "", "", "", now[:10], 0, "0", 0, "0", 0, "0", "", 0, "", "", "", set_num, "", "", "", "", event_code, "", "", 1, "", ""]
        cursor.execute(f"INSERT INTO dbo.Event ({columns}) VALUES ({','.join('?' for _ in values)})", values)

    @staticmethod
    def _insert_event_traveler(cursor, account_id: str, set_num: int, customer_name: str, phone: str, values: dict[str, str], now: str) -> None:
        cursor.execute("SELECT 1 FROM dbo.Event_Traveler WHERE ID=? AND Num=?", account_id, set_num)
        if cursor.fetchone():
            return
        columns = "ID,Num,ETGuBun,Name_K,Sex,Name_EF,Name_EL,BirthDay,CabinType,Tel,EMail,PassPortNo,PassPortEDate,Save_Date,Modi_Date,UsrID"
        row = [account_id, set_num, "기본", customer_name, "", values["영문성"], values["영문이름"], values["생년월일"], values["객실종류"], phone, "", values["여권번호"], values["여권만료일"], now, now, "Admin"]
        cursor.execute(f"INSERT INTO dbo.Event_Traveler ({columns}) VALUES ({','.join('?' for _ in row)})", row)

    @staticmethod
    def _upsert_event_expenses(cursor, account_id: str, set_num: int, expense_values: list[object], now: str) -> None:
        amount_columns = ["PortCh", "FuelSurCha", "ExcRateAddCha", "IntraTip", "RoomUpCha", "AirUpCha", "VisaCha", "EtcCha1"]
        numeric_columns = [
            "PrePay", "SupPay", "PortCh", "FuelSurCha", "ExcRateAddCha", "IntraTip",
            "RoomUpYN", "RoomUpCha", "AirUpYN", "AirUpCha", "VisaCha", "RoomDCCha",
            "AirAddCha", "SpeAddCha", "AirDCCha", "OneTimePay", "Penalty", "EtcCha1",
            "EtcCha2", "EtcCha3", "EtcCha4", "EtcCha5",
        ]
        amount_values = [0 if value is None else value for value in expense_values]
        values_by_column = dict(zip(amount_columns, amount_values))
        values_by_column["RoomUpYN"] = 1 if float(values_by_column["RoomUpCha"] or 0) > 0 else 0
        values_by_column["AirUpYN"] = 1 if float(values_by_column["AirUpCha"] or 0) > 0 else 0
        write_columns = [*amount_columns, "RoomUpYN", "AirUpYN"]
        write_values = [values_by_column[column] for column in write_columns]
        cursor.execute("SELECT 1 FROM dbo.Event_Expenses WHERE ID=? AND Num=?", account_id, set_num)
        if cursor.fetchone():
            unmapped_columns = [column for column in numeric_columns if column not in write_columns]
            set_clauses = [*(f"{column}=?" for column in write_columns)]
            set_clauses.extend(f"{column}=ISNULL({column},0)" for column in unmapped_columns)
            set_clauses.extend(["Modi_Date=?", "UsrID='Admin'"])
            cursor.execute(
                f"UPDATE dbo.Event_Expenses SET {','.join(set_clauses)} WHERE ID=? AND Num=?",
                [*write_values, now, account_id, set_num],
            )
            return
        numeric_values = {column: 0 for column in numeric_columns}
        numeric_values.update(values_by_column)
        insert_columns = "ID,Num," + ",".join(numeric_columns) + ",Save_Date,Modi_Date,UsrID"
        row = [account_id, set_num, *(numeric_values[column] for column in numeric_columns), now, now, "Admin"]
        cursor.execute(f"INSERT INTO dbo.Event_Expenses ({insert_columns}) VALUES ({','.join('?' for _ in row)})", row)

    @staticmethod
    def _update_member_status(cursor, account_id: str, departure: str, member_status: str, now: str) -> None:
        cursor.execute("SELECT memtype FROM dbo.member WHERE ID=?", account_id)
        found = cursor.fetchone()
        if not found:
            return
        old_memtype = str(found[0] or "")
        cursor.execute("UPDATE dbo.member SET memtype=? WHERE ID=?", member_status, account_id)
        cursor.execute("SELECT 1 FROM dbo.TransMType WHERE ID=? AND EDate=?", account_id, departure)
        if cursor.fetchone():
            cursor.execute("UPDATE dbo.TransMType SET EMemType=? WHERE ID=? AND EDate=?", member_status, account_id, departure)
        else:
            cursor.execute("INSERT INTO dbo.TransMType (EDate,ID,EMemType) VALUES (?,?,?)", departure, account_id, member_status)
        memo = f"여행 등록처리 완료 : {old_memtype} → {member_status}"
        cursor.execute("INSERT INTO dbo.M_Memo (Note_Date,Note_time,Charge_ID,ID,TMemo,Save_Date,Modi_Date,TelCall_Gu,TelCall_ErrGu,TelCall_Mem,Memo_GuBun,ProDate,M_Submit,M_PayNum,etc_str1) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", now[:10], now[11:], "Admin", account_id, memo, now, now, "여행등록처리", "기타", "", "0", "", "OTHER", 0, 0)

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
