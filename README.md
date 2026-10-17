# 💳 카드 명세서 분석기

신용카드 월별 명세서 엑셀 파일을 올리면 사용처를 **식비/카페 · 교통/차량 · 쇼핑 · 주거/통신 · 의료/건강 · 문화/여가 · 기타**로 자동 분류하고, 도넛 차트 대시보드와 정리된 엑셀 파일을 제공하는 Streamlit 웹 앱입니다.

## 주요 기능

| 기능 | 설명 |
|---|---|
| 엑셀 업로드 | `.xlsx`, `.xls` 지원 (카드사의 HTML/XML 형식 '가짜 xls'도 그대로 업로드). 헤더 행 자동 감지 |
| 열 자동 인식 | 이용일자·가맹점·결제금액 열을 열 이름으로 자동 인식 (예: 신한카드 `거래일자`·`이용가맹점`·`결제 금액`). 결제금액 0원 행은 제외 |
| 금액 전처리 | `12,300원`, `₩5,000`, `(1,000)`, `-3,000` 등 안전하게 숫자로 변환. 합계/소계 행 자동 제외 |
| 자동 분류 | 국내 주요 가맹점 키워드 사전 기반 1차 분류 + **내가 추가한 키워드 우선 적용** |
| 내 분류·키워드 | 사이드바에서 기본 분류에 키워드 추가/삭제, 새 분류를 색상과 함께 생성. 대시보드에 즉시 반영 |
| 수동 편집 | 번호·이용일자·가맹점·결제금액·분류 고정 표에서 '분류' 칸만 드롭다운으로 수정 (나머지는 잠금). '기타만 보기' 필터 |
| 대시보드 | 총 지출, 이용 건수, 최대 지출 분류, 도넛 차트, 카테고리별 요약 표 (**분류 행 클릭 시 상세 내역 표시**) |
| 옵션 | 취소·환불(마이너스) 제외 토글, 여러 달이 섞인 경우 월 필터 |
| 다운로드 | `전체내역` + `카테고리별합계` 두 시트로 구성된 `.xlsx` |

업로드한 파일은 서버 디스크에 저장하지 않고 메모리에서만 처리합니다.

## 프로젝트 구조

```
Credit_Card_Tracker/
├── app.py            # 메인 애플리케이션
├── requirements.txt  # 패키지 목록
├── user_keywords.json # 내가 추가한 분류/키워드 (앱에서 처음 추가할 때 자동 생성)
└── README.md
```

## 로컬 실행 방법

Python 3.9 이상이 필요합니다.

### Windows (PowerShell)

```powershell
cd C:\Learning_python\Credit_Card_Tracker
python -m venv .venv
.\.venv\Scripts\Activate.ps1          # 실행 정책 오류 시: Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
pip install -r requirements.txt
streamlit run app.py
```

### macOS / Linux

```bash
cd Credit_Card_Tracker
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

브라우저가 자동으로 `http://localhost:8501` 을 엽니다.

> 같은 와이파이에 있는 지인에게 잠깐 보여주려면 터미널에 표시되는 `Network URL` (예: `http://192.168.0.10:8501`)로 접속하면 됩니다.

## Streamlit Community Cloud 무료 배포

지인들이 언제든 브라우저로 접속할 수 있도록 무료로 배포할 수 있습니다.

1. **GitHub 저장소 만들기** — github.com 에서 새 저장소(Public 또는 Private)를 만듭니다.
2. **코드 올리기**
   ```bash
   git remote add origin https://github.com/<아이디>/<저장소명>.git
   git push -u origin main
   ```
   `app.py`, `requirements.txt` 가 저장소 최상위에 있어야 합니다. (`.venv` 폴더는 `.gitignore`로 제외됨)
3. **Streamlit Cloud 접속** — <https://share.streamlit.io> 에서 GitHub 계정으로 로그인합니다.
4. **앱 생성** — `Create app` → `Deploy a public app from GitHub` 선택 후
   - Repository: `<아이디>/<저장소명>`
   - Branch: `main`
   - Main file path: `app.py`
   - (선택) App URL: 원하는 주소 지정 (예: `my-card-tracker`)
   - `Advanced settings`에서 Python 버전 선택 가능 (3.11~3.12 권장)
5. **Deploy** 클릭 → 1~3분 후 `https://<앱이름>.streamlit.app` 주소가 생성됩니다. 이 링크를 지인들에게 공유하세요.

**배포 팁**
- 코드를 GitHub에 `push`하면 앱이 자동으로 재배포됩니다.
- 무료 플랜은 일정 기간 접속이 없으면 앱이 잠자기 상태가 되며, 접속 시 버튼을 눌러 깨울 수 있습니다.
- 접근을 지인으로 제한하고 싶다면 앱 `Settings → Sharing`에서 이메일로 초대할 수 있습니다.
- 명세서는 개인 금융정보입니다. 공개 앱으로 배포하더라도 업로드 데이터는 각 사용자 세션 메모리에만 머무르며 다른 사용자와 공유되지 않습니다.

## 분류 키워드 추가

앱 왼쪽 사이드바의 **🏷️ 분류 키워드**에서 코드 수정 없이 관리합니다.

- **기본 분류에 키워드 추가**: 분류를 펼쳐 '내 키워드'에 입력 후 Enter (✕로 삭제)
- **새 분류 만들기**: 이름·색상·키워드(쉼표 구분)를 입력하고 '분류 추가'. 만든 분류는 색상 변경·삭제 가능
- 내 키워드는 기본 키워드보다 **먼저** 적용되며, 긴 키워드가 우선합니다. 대소문자·공백은 무시하고 **부분 일치**로 비교합니다.
- 표에서 직접 바꾼 분류는 키워드보다 우선합니다.
- 설정은 `user_keywords.json`에 저장되어 앱을 다시 켜도 유지됩니다. (Streamlit Cloud에서는 앱이 재시작되면 초기화되고, 접속한 모든 사람이 같은 설정을 공유합니다.)

기본 키워드 자체를 바꾸려면 `app.py` 상단의 `CATEGORY_KEYWORDS`를 수정하세요.

## 문제 해결

| 증상 | 해결 |
|---|---|
| 엑셀에서 열 때 "파일 형식 및 확장자가 일치하지 않습니다" 경고 | 카드사 `.xls`가 실제로는 HTML/XML 파일이라 생기는 경고입니다. 앱이 파일 내용을 보고 자동 변환하므로 **다시 저장하지 말고 그대로 업로드**하면 됩니다 |
| `.xls` 파일이 열리지 않음 | `pip install -r requirements.txt` 로 `xlrd`, `lxml` 설치 확인 |
| 암호가 걸린 파일 | 엑셀에서 암호를 해제한 뒤 다시 저장해서 업로드 |
| '가맹점/결제금액 열을 찾지 못했어요' | '원본 데이터 미리보기'에서 열 이름을 확인하고, `app.py`의 `COLUMN_HINTS`에 해당 열 이름을 추가 |
