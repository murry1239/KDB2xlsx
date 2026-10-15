# KDB2xlsx — 원시자료 정리봇 alpha_0.3.1

KAERS 이상사례 원시자료 txt 13개(README, DEMO, HIST_E, PARENT, EVENT, TEST, DRUG, DRUG1~3, DRUG_EVENT, ASSESSMENT, GROUP)를
읽어 정리된 엑셀 파일(_01, _02, _03, _10, _20, _30)을 만드는 **2단계 화면 프로그램**입니다.

## 작업 순서

| 단계 | 화면 탭 | 입력 | 결과 |
|---|---|---|---|
| A | **A. 01~03 만들기** | txt 폴더, 제품명, 의약품명 | `_01`, `_02`, `_03` |
| 사용자 | (엑셀) | `_03` 파일 EVENT 시트 **PT_한국어·SOC_한국어** 입력 후 저장 | |
| B | **B. 10~30 만들기** | 입력을 마친 `_03` 파일 | `_10`, `_20`, `_30` |
| 사용자 | (엑셀) | `_30` 파일을 확인하며 **LLT_한국어·반영여부** 직접 입력 (제목 행만 만들어 둠) | |

- A 실행이 끝나면 `_03` 파일을 바로 열 수 있고, B 탭에 파일·제품명·의약품명이 자동으로 채워집니다.
- B 탭에서 `_03` 파일을 직접 고르면 파일 이름에서 제품명을, DRUG 시트에서 의약품명을 자동으로 읽습니다.
- B 실행 시 PT_한국어·SOC_한국어에 빈칸이 있으면 진행 기록에 행 수를 표시합니다(작업은 계속 진행).
- 결과 저장 폴더 기본값: A는 txt 폴더가 `files`이면 그 상위(접수번호) 폴더, B는 `_03` 파일이 있는 폴더.

## 설치와 실행

### 방법 1 — exe (Python 없는 PC)
`원시자료_정리봇_alpha_0.3.1.exe`를 더블클릭합니다. 명령 프롬프트 창 없이 프로그램 화면만 열립니다.

### 방법 2 — 소스 폴더 (Python 3.11 이상)
1. 처음 한 번 `INSTALL.bat` 실행 (가상환경 `.venv`와 openpyxl 설치, 인터넷 필요)
2. 이후에는 `RUN.bat`만 실행
3. 문제가 생기면 `DIAGNOSE.bat` 결과를 전달

## 오류와 로그

- 화면 상단 **로그 열기**: 이번 실행의 로그 파일, **로그 폴더**: 지난 로그 모음
- 로그 위치: `%LOCALAPPDATA%\원시자료정리봇\logs\rawbot_YYYYMMDD_HHMMSS.log`
- 오류가 나면 창에 원인과 로그 위치를 표시합니다. 로그 파일을 함께 전달하면 원인을 찾기 쉽습니다.
- 같은 이름의 엑셀 파일이 열려 있으면 저장할 수 없으니 닫고 다시 실행하세요.
- 실행 중 자료를 외부 AI나 서버로 보내지 않습니다.

## 파일 구성

| 파일 | 내용 |
|---|---|
| `app.py` | 화면(tkinter), 명령줄 실행 |
| `engine.py` | 처리 엔진 (txt 읽기, 열 추가, 코드 변환, 정보 불러오기, 저장) |
| `tests/test_engine.py` | 가상 자료로 A → 입력 → B 전체 흐름 검증 |
| `INSTALL.bat` / `RUN.bat` / `DIAGNOSE.bat` / `TEST.bat` | 설치·실행·진단·시험 |
| `docs/` | 작업 명세와 claude.ai 채팅 작업기록 |

## 개발과 검증

```bash
py -m unittest discover -s tests -v
py app.py --step A --txt "...\files" --product 제품명 --drug-name 의약품명
py app.py --step B --xlsx "...\원시자료_제품명_03_YYYYMMDD.xlsx" --sheet-prefix 시트명
py -m PyInstaller --onefile --windowed --name 원시자료_정리봇_alpha_0.3.1 app.py
```

변경 이력은 [CHANGELOG.md](CHANGELOG.md)에 있습니다.

## 주의
원시자료(txt)와 결과 엑셀은 민감정보이므로 `.gitignore`로 저장소에서 제외합니다. exe(`dist/`)도 저장소에 올리지 않습니다.
