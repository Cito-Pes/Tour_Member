# 여행고객 CRM 등록 (항차별)

여행사에서 전달받은 항차별 Excel 고객자료를 조회하고, 회원코드에 연결된 회원번호를 검토한 뒤 고객관리 DB에 일괄 반영하는 Windows 전용 PySide6 프로그램이다.

## 버전

- 프로그램 버전: 1.0
- Python: 3.13.x
- UI: PySide6
- 배포: PyInstaller Windows 단일 EXE
- 문자 인코딩: UTF-8
- 기본 글꼴: D2Coding

## 주요 기능

- `.xlsx`, `.xls` 여행고객 자료 조회
- 첫 번째 안내행에서 행사코드, 여행지역, 출발일, 도착일 추출
- 원본 Excel 파일을 변경하지 않고 화면에서만 데이터 가공
- DB에 등록된 항차와 회원번호 처리 여부 조회
- 회원코드 기준 회원번호 1~4 자동 할당
- 검토 단계에서 기초자료 저장
- 적용 단계에서 행사, 여행자, 비용 및 회원상태 일괄 반영
- 화면 데이터를 별도 Excel 파일로 다운로드
- 처리현황과 오류 발생 단계를 로그로 표시

## 데이터 처리 구조

1. `② 조회`에서 Excel 자료를 화면에 표시한다.
2. `③ 검토`에서 자료를 확인하고 `Tour_Member_M`에 기초자료를 저장한다.
3. 회원코드에 연결된 회원번호를 조회해 회원번호 1~4에 할당한다.
4. `④ 적용`에서 `Tour_Member_Account`에 할당 결과를 저장한다.
5. 같은 트랜잭션에서 `ET_Goods`, `Event`, `Event_Traveler`, `Event_Expenses`, `member`, `TransMType`, `M_Memo`를 반영한다.
6. 적용 중 하나라도 실패하면 적용 단계 전체를 롤백한다. 검토 단계에서 이미 저장된 기초자료는 유지된다.

## 특수 처리

- 예약상태가 `TOUR LEADER`이면 회원코드에 `TOUR LEADER`를 입력하고 예약상태는 비운다.
- `TOUR LEADER`는 회원명이 없어도 기초자료 검증을 통과한다.
- 주민번호 뒷자리는 Excel에 값이 있어도 DB에 저장하지 않는다.
- 메모에 `노쇼`가 포함되면 여행상태를 `노쇼`로 표시한다.
- 날짜는 `YYYY-MM-DD` 형식으로 저장한다.
- 화면의 숫자는 오른쪽 정렬하고 3자리 쉼표로 표시한다.
- 원본 Excel 파일은 수정하지 않는다.

## 비용 매핑

| Excel 컬럼 | Event_Expenses 컬럼 |
|---|---|
| 항만세 | PortCh |
| 유류할증료 | FuelSurCha |
| 환율추가금 | ExcRateAddCha |
| 선내팁 | IntraTip |
| 객실추가 | RoomUpCha |
| 항공추가 | AirUpCha |
| 비자(금액) | VisaCha |
| 추가금액 | EtcCha1 |

`Event_Expenses`의 숫자형 컬럼은 빈 값일 때 `0`으로 저장한다. `RoomUpCha`가 0보다 크면 `RoomUpYN=1`, `AirUpCha`가 0보다 크면 `AirUpYN=1`로 저장하고, 그렇지 않으면 각각 `0`으로 저장한다.

## 프로젝트 실행

```powershell
.\.venv\Scripts\Activate.ps1
python main.py
```

의존성은 `requirements.txt`를 사용한다.

```powershell
python -m pip install -r requirements.txt
```

## EXE 빌드

```powershell
.\.venv\Scripts\python.exe -m PyInstaller `
  --noconfirm --clean --onefile --windowed `
  --name "여행고객 CRM 등록 (항차별)" `
  --icon "img\ht.ico" `
  --version-file "version_info.txt" `
  --add-data "config.txt;." `
  "main.py"
```

결과 파일은 `dist\여행고객 CRM 등록 (항차별).exe`로 생성된다. `Config_DB.db`는 EXE에 포함하지 않으며, 최초 실행 시 내장된 `config.txt` 정보를 이용해 EXE 옆의 `DB` 폴더에 자동 다운로드한다.

상세 사용방법은 [USER_MANUAL.md](USER_MANUAL.md)를 참고한다.
