# KDB2xlsx — 원시자료 정리봇

KAERS 이상사례 원시자료 txt 13개(README, DEMO, HIST_E, PARENT, EVENT, TEST, DRUG, DRUG1~3, DRUG_EVENT, ASSESSMENT, GROUP)를
읽어 정리된 엑셀 파일(_01, _02, _03, _10, _20, _30)을 만드는 파이썬 스크립트입니다.

## 필요 환경
- Python 3.11
- `pip install openpyxl`

## 실행
```bash
python 원시자료_정리봇.py
```
화면 질문에 따라 작업 번호, txt 폴더, 제품명 등을 입력합니다.

| 번호 | 작업 | 결과 |
|---|---|---|
| 1 | 전체 실행 | _01 ~ _30 |
| 2 | 1~3단계 | _01, _02, _03 (이후 EVENT 시트에 PT/SOC 직접 입력) |
| 3 | 4~6단계 | PT/SOC 입력한 _03 → _10, _20, _30 |

## exe 만들기 (Python 없는 PC용)
```bash
py -m pip install pyinstaller
py -m PyInstaller --onefile --console --name 원시자료_정리봇_alpha_0.1 원시자료_정리봇.py
```
`dist/원시자료_정리봇_alpha_0.1.exe` 가 생성됩니다.

## 버전 이력
### alpha_0.1 (2026-10-14)
- 실제 원시자료(글리팜정4mg, 접수번호 2610B0048, 보고 15건)로 첫 실행 검증
  - 의심 6행 / 병용 15행: txt에서 직접 센 값과 일치
  - DRUG_EVENT 190행, #N/A 없음
- 수정: 앞자리 0이 있는 코드(예: 연령단위 `00103`)가 변환표와 맞지 않아 빈칸이 되던 문제
- 수정: 한글 Windows(cp949)에서 출력을 파일로 돌리면 `✔` 문자 때문에 멈추던 문제
- exe 빌드 추가

### 초기 버전
- claude.ai 채팅에서 작성 (그리타존정15mg 수작업 결과와 셀 단위로 비교 검증)

## 문서
- [docs/명세_및_클로드챗_작업기록.txt](docs/명세_및_클로드챗_작업기록.txt): 작업 명세와 claude.ai 채팅에서 진행한 내용

## 주의
원시자료(txt)와 결과 엑셀은 민감정보이므로 `.gitignore`로 저장소에서 제외합니다.
