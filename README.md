# 우신 업무관리 시스템

HTML/JS 프론트엔드 + Google Apps Script API + Google Sheets 데이터

---

## 폴더 구조

```
wooshin-work/
├── index.html              ← 메인 파일 (HTML + JS 전체)
├── README.md
└── .github/
    └── workflows/
        └── deploy.yml      ← GitHub Pages 자동 배포
```

---

## 1단계 — Google Sheets 생성

1. [Google Sheets](https://sheets.google.com) 접속 → 새 스프레드시트 생성
2. 이름: `우신업무관리_DB`
3. URL에서 스프레드시트 ID 복사
   ```
   https://docs.google.com/spreadsheets/d/【이 부분】/edit
   ```

---

## 2단계 — Apps Script 배포

1. Sheets 메뉴 → **확장 프로그램 → Apps Script**
2. `Code.gs` 내용을 전체 붙여넣기
3. 상단 `SS_ID = ''` 에 스프레드시트 ID 입력
   ```js
   var SS_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms'; // 예시
   ```
4. **저장 (Ctrl+S)**
5. 함수 선택 드롭다운에서 `setupAllSheets` 선택 → **실행** (최초 1회)
   - 권한 허용 팝업 → 허용
   - Sheets에 자동으로 시트 8개 생성됨
6. 메뉴 → **배포 → 새 배포**
   - 유형: **웹 앱**
   - 설명: `v1`
   - 실행 계정: **나**
   - 액세스 권한: **모든 사용자**
   - **배포** 클릭
7. 웹 앱 URL 복사 (형식: `https://script.google.com/macros/s/XXXX/exec`)

---

## 3단계 — GitHub 저장소 설정

1. GitHub 접속 → **New repository**
   - 이름: `wooshin-work`
   - Public (GitHub Pages 무료 사용)
   - README 생성 체크 해제
2. `index.html`, `README.md`, `.github/workflows/deploy.yml` 업로드
3. 저장소 설정 → **Pages** → Source: **GitHub Actions** 선택
4. `main` 브랜치에 push → 자동 배포 시작

---

## 4단계 — GAS URL 연결

1. 배포된 GitHub Pages URL 접속
2. 로그인 → **관리자** → **설정** 탭
3. GAS URL 입력란에 2단계에서 복사한 웹 앱 URL 붙여넣기
4. **저장 & 연결 테스트** 클릭 → ✅ 연결 성공

---

## 시트 구조

| 시트명 | 주요 컬럼 |
|--------|-----------|
| 직원 | id, pw, name, role, team, rank... |
| 실적 | empId, empName, date, attendStatus, rows(JSON) |
| 연차 | empId, empName, date, type |
| IP그룹 | id, type, name, prefix, entries(JSON) |
| 변경로그 | dt, by, desc |
| 연차공지 | notice |
| 자리배치 | id, label, empId |
| 설정 | key, value |

---

## 코드 수정 방법

1. `index.html` 수정
2. GitHub에 push → 자동으로 Pages 재배포 (약 1~2분)

```bash
git add index.html
git commit -m "기능 수정: 내용"
git push origin main
```

---

## 주의사항

- GAS 웹 앱 코드 수정 시 반드시 **새 배포** 또는 **배포 관리 → 버전 업데이트** 필요
- 기존 배포 URL은 유지됨 (HTML에서 재설정 불필요)
- 직원 비밀번호는 Sheets에 평문 저장됨 — 민감 환경에서는 해시 처리 권장
