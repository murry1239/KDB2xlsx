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

## 문서
- [docs/명세_및_클로드챗_작업기록.txt](docs/명세_및_클로드챗_작업기록.txt): 작업 명세와 claude.ai 채팅에서 진행한 내용

## 주의
원시자료(txt)와 결과 엑셀은 민감정보이므로 `.gitignore`로 저장소에서 제외합니다.
