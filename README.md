# 김수철 교장선생님께 보내는 마음

광평중학교 이임 롤링페이퍼. 학생·교직원·학부모가 이름과 함께 편지를 남기고, 교장선생님이 읽는 웹앱.

- `index.html` — 프론트엔드(단일 HTML, GitHub Pages 배포)
- `gas/Code.gs` — 백엔드 API(Google Apps Script + Google Sheets)
- `시안/` — 화면 시안 모음

## 주소
- 편지 쓰기·보기: https://thanks.kimju.kr/
- 관리자(편지 내리기): 주소 뒤에 `?admin=1`

## 운영 메모
- 편지는 Google Sheets `편지` 시트에 쌓인다.
- 관리 비밀번호는 스크립트 속성 `ADMIN_PIN`에 있다(기본 1234 — 공개 전 변경).
- 백엔드 수정 시 새 배포를 만들지 말고 기존 배포를 새 버전으로 갱신한다.
