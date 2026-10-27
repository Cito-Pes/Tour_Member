# -*- coding: utf-8 -*-
"""Repair historical Event_Expenses values created before per-account splitting.

Default mode is preview-only. Use --apply only after reviewing the preview.
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from DB_conn import connect


AMOUNT_COLUMNS = [
    "PortCh",
    "FuelSurCha",
    "ExcRateAddCha",
    "IntraTip",
    "RoomUpCha",
    "AirUpCha",
    "VisaCha",
    "EtcCha1",
]
SOURCE_COLUMNS = [
    "PortTax",
    "FuelSurcharge",
    "ExchangeRateExtra",
    "OnboardTip",
    "RoomExtra",
    "AirlineExtra",
    "VisaAmount",
    "ExtraAmount",
]
NUMERIC_COLUMNS = [
    "PrePay", "SupPay", "PortCh", "FuelSurCha", "ExcRateAddCha", "IntraTip",
    "RoomUpYN", "RoomUpCha", "AirUpYN", "AirUpCha", "VisaCha", "RoomDCCha",
    "AirAddCha", "SpeAddCha", "AirDCCha", "OneTimePay", "Penalty", "EtcCha1",
    "EtcCha2", "EtcCha3", "EtcCha4", "EtcCha5",
]
UNMAPPED_NUMERIC_COLUMNS = [
    column for column in NUMERIC_COLUMNS
    if column not in {*AMOUNT_COLUMNS, "RoomUpYN", "AirUpYN"}
]


def trunc_divide(amount: int, divisor: int) -> int:
    """Divide an integer toward zero without introducing a decimal value."""
    sign = -1 if amount < 0 else 1
    return sign * (abs(amount) // divisor)


def split_amount(amount: int, divisor: int, position: int, count: int) -> int:
    base = trunc_divide(amount, divisor)
    if position == count:
        return amount - (base * (count - 1))
    return base


def load_rows(cursor) -> list[dict[str, object]]:
    cursor.execute(
        """
        SELECT
            a.TourMemberAccountSeqNo,
            a.TourMemberSeqNo,
            a.AccountID,
            a.AllocationOrder,
            e.Num,
            m.EventCode,
            m.PortTax,
            m.FuelSurcharge,
            m.ExchangeRateExtra,
            m.OnboardTip,
            m.RoomExtra,
            m.AirlineExtra,
            m.VisaAmount,
            m.ExtraAmount,
            x.PortCh,
            x.FuelSurCha,
            x.ExcRateAddCha,
            x.IntraTip,
            x.RoomUpCha,
            x.AirUpCha,
            x.VisaCha,
            x.EtcCha1
        FROM dbo.Tour_Member_Account AS a
        INNER JOIN dbo.Tour_Member_M AS m
            ON m.SeqNo = a.TourMemberSeqNo
        INNER JOIN dbo.Event AS e
            ON e.ID = a.AccountID
        INNER JOIN dbo.Event_Expenses AS x
            ON x.ID = e.ID
           AND x.Num = e.Num
        WHERE a.AllocationStatus = 'ALLOCATED'
        ORDER BY a.TourMemberSeqNo, a.AllocationOrder, a.TourMemberAccountSeqNo
        """
    )
    rows: list[dict[str, object]] = []
    for row in cursor.fetchall():
        rows.append(
            {
                "account_seq": int(row[0]),
                "parent_seq": int(row[1]),
                "account_id": str(row[2]),
                "allocation_order": int(row[3]),
                "num": int(row[4]),
                "event_code": str(row[5] or ""),
                "source": [int(row[index] or 0) for index in range(6, 14)],
                "current": [int(row[index] or 0) for index in range(14, 22)],
            }
        )
    return rows


def build_stage(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    groups: dict[int, list[dict[str, object]]] = defaultdict(list)
    for row in rows:
        groups[int(row["parent_seq"])].append(row)

    stage: list[dict[str, object]] = []
    for parent_seq, group in groups.items():
        orders = [int(row["allocation_order"]) for row in group]
        if len(orders) != len(set(orders)):
            raise RuntimeError(f"Tour_Member_M.SeqNo={parent_seq}: AllocationOrder가 중복돼.")
        group.sort(key=lambda row: int(row["allocation_order"]))
        account_count = len(group)
        if account_count <= 0:
            raise RuntimeError(f"Tour_Member_M.SeqNo={parent_seq}: 할당 계정이 없어.")
        source = list(group[0]["source"])
        for row in group[1:]:
            if list(row["source"]) != source:
                raise RuntimeError(f"Tour_Member_M.SeqNo={parent_seq}: 원본 금액이 계정별로 다르게 조회돼.")
        for position, row in enumerate(group, start=1):
            amounts = [
                split_amount(amount, account_count, position, account_count)
                for amount in source
            ]
            stage.append(
                {
                    **row,
                    "account_count": account_count,
                    "expected": amounts,
                    "room_up_yn": 1 if amounts[4] > 0 else 0,
                    "air_up_yn": 1 if amounts[5] > 0 else 0,
                }
            )
        for index, amount in enumerate(source):
            if sum(int(row["expected"][index]) for row in stage if int(row["parent_seq"]) == parent_seq) != amount:
                raise RuntimeError(f"Tour_Member_M.SeqNo={parent_seq}: 원본 금액 합계 검증에 실패했어.")
    return stage


def create_backup(cursor, backup_name: str) -> int:
    cursor.execute(
        f"""
        SELECT DISTINCT x.*
        INTO dbo.[{backup_name}]
        FROM dbo.Event_Expenses AS x
        INNER JOIN dbo.Event AS e
            ON e.ID = x.ID
           AND e.Num = x.Num
        INNER JOIN dbo.Tour_Member_Account AS a
            ON a.AccountID = e.ID
           AND a.AllocationStatus = 'ALLOCATED'
        INNER JOIN dbo.Tour_Member_M AS m
            ON m.SeqNo = a.TourMemberSeqNo
        """
    )
    cursor.execute(f"SELECT COUNT(*) FROM dbo.[{backup_name}]")
    return int(cursor.fetchone()[0])


def create_stage_table(cursor) -> None:
    cursor.execute(
        """
        CREATE TABLE #EventExpensesRepair (
            AccountSeqNo INT NOT NULL,
            ParentSeqNo INT NOT NULL,
            AccountID VARCHAR(100) NOT NULL,
            Num INT NOT NULL,
            AllocationOrder INT NOT NULL,
            AccountCount INT NOT NULL,
            PortCh INT NOT NULL,
            FuelSurCha INT NOT NULL,
            ExcRateAddCha INT NOT NULL,
            IntraTip INT NOT NULL,
            RoomUpCha INT NOT NULL,
            AirUpCha INT NOT NULL,
            VisaCha INT NOT NULL,
            EtcCha1 INT NOT NULL,
            RoomUpYN SMALLINT NOT NULL,
            AirUpYN SMALLINT NOT NULL
        )
        """
    )


def insert_stage(cursor, stage: list[dict[str, object]]) -> None:
    cursor.executemany(
        """
        INSERT INTO #EventExpensesRepair
        (AccountSeqNo,ParentSeqNo,AccountID,Num,AllocationOrder,AccountCount,
         PortCh,FuelSurCha,ExcRateAddCha,IntraTip,RoomUpCha,AirUpCha,VisaCha,EtcCha1,
         RoomUpYN,AirUpYN)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,
        [
            [
                row["account_seq"], row["parent_seq"], row["account_id"], row["num"],
                row["allocation_order"], row["account_count"], *row["expected"],
                row["room_up_yn"], row["air_up_yn"],
            ]
            for row in stage
        ],
    )


def apply_rows(cursor, stage: list[dict[str, object]]) -> None:
    set_clauses = [
        "PortCh=?", "FuelSurCha=?", "ExcRateAddCha=?", "IntraTip=?",
        "RoomUpCha=?", "AirUpCha=?", "VisaCha=?", "EtcCha1=?",
        "RoomUpYN=?", "AirUpYN=?",
    ]
    set_clauses.extend(f"{column}=ISNULL({column},0)" for column in UNMAPPED_NUMERIC_COLUMNS)
    sql = f"UPDATE dbo.Event_Expenses SET {', '.join(set_clauses)} WHERE ID=? AND Num=?"
    for row in stage:
        cursor.execute(
            sql,
            [*row["expected"], row["room_up_yn"], row["air_up_yn"], row["account_id"], row["num"]],
        )


def verify_rows(cursor, stage: list[dict[str, object]]) -> int:
    mismatch_count = 0
    cursor.execute(
        """
        SELECT x.PortCh,x.FuelSurCha,x.ExcRateAddCha,x.IntraTip,
               x.RoomUpCha,x.AirUpCha,x.VisaCha,x.EtcCha1,
               x.RoomUpYN,x.AirUpYN
        FROM dbo.Event_Expenses AS x
        INNER JOIN #EventExpensesRepair AS s
            ON s.AccountID=x.ID AND s.Num=x.Num
        ORDER BY s.AccountSeqNo
        """
    )
    for actual, expected in zip(cursor.fetchall(), stage):
        expected_values = [*expected["expected"], expected["room_up_yn"], expected["air_up_yn"]]
        if list(actual) != expected_values:
            mismatch_count += 1
    return mismatch_count


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Event_Expenses historical repair")
    parser.add_argument("--apply", action="store_true", help="create backup and apply the validated repair")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    with connect() as connection:
        cursor = connection.cursor()
        rows = load_rows(cursor)
    if not rows:
        print("보정 대상이 없어.")
        return 0
    stage = build_stage(rows)
    parent_count = len({int(row["parent_seq"]) for row in stage})
    account_count_distribution: dict[int, int] = defaultdict(int)
    for row in stage:
        account_count_distribution[int(row["account_count"])] += 1
    print(f"대상 부모 행: {parent_count}건")
    print(f"대상 계정 행: {len(stage)}건")
    print("계정 수별 행:", ", ".join(f"{count}계정 {rows_count}건" for count, rows_count in sorted(account_count_distribution.items())))
    for row in stage[:10]:
        print(
            f"미리보기: {row['event_code']} / 부모 {row['parent_seq']} / "
            f"계정 {row['account_id']} / {row['allocation_order']}/{row['account_count']} / "
            f"환율추가금 {row['source'][2]} -> {row['expected'][2]}"
        )
    if not args.apply:
        print("미리보기만 수행했어. DB 변경 없음.")
        return 0

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    backup_name = f"Event_Expenses_RepairBackup_{timestamp}"
    with connect() as connection:
        cursor = connection.cursor()
        backup_count = create_backup(cursor, backup_name)
        connection.commit()
    print(f"백업 완료: dbo.{backup_name} ({backup_count}건)")

    with connect() as connection:
        cursor = connection.cursor()
        try:
            create_stage_table(cursor)
            insert_stage(cursor, stage)
            apply_rows(cursor, stage)
            mismatch_count = verify_rows(cursor, stage)
            if mismatch_count:
                raise RuntimeError(f"보정 후 값 검증 실패: {mismatch_count}건")
            connection.commit()
        except Exception:
            connection.rollback()
            raise
    print(f"보정 적용 완료: {len(stage)}건")
    print(f"백업 테이블: dbo.{backup_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
