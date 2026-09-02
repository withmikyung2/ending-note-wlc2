# GYA_AMS (GYA Academic Management System)

캄보디아 뚬놉(Tom Nop) 지역 GYA Center와 7개 Community Learning Room을 위한
학사관리(출석·학생·교사·평가·활동) 시작용 시스템입니다.

이 폴더는 기존 `wlc-note` 사이트(엔딩노트 강의 사이트)와 별개의 독립 프로젝트입니다.
기존 폴더/파일은 수정하지 않았습니다.

## 배경

- 근거 문서: 「GYA·GWC 뚬놉 지역 교육혁신 네트워크 구축사업」 제안서 (2026.09 Draft)
- GYA는 14년간 캄보디아 농촌지역 아동·청소년의 학습권을 지원해 왔고,
  뚬놉 지역 7개 자율 공부방을 하나의 교육지원체계로 연결하는 것이 이 사업의 목표입니다.

## 구성

- `index.html` — 프로젝트 소개 및 안내 페이지 (정적 사이트)
- `docs/data-model.md` — Google Sheets 데이터 구조(탭·필드) 명세
- `docs/forms-spec.md` — Google Forms로 만들어야 할 입력폼 필드 명세
- `scripts/build_workbook.py` — 데이터시트(xlsx) 초안을 생성하는 스크립트 (`python3 build_workbook.py`,
  `openpyxl` 필요). 구조를 수정할 때는 이 스크립트를 고쳐서 다시 생성합니다.
- 실제 데이터가 들어간 시트 파일(xlsx/Google Sheets)은 이 저장소에 포함하지 않고 Google Drive에
  별도로 둡니다 (개인정보·아동정보가 포함되므로 공개 저장소에 올리지 않습니다).

## 왜 로우코드(Google Sheets/Forms)인가

- 비용 0원, 현지 Local Teacher가 이미 구글폼/구글시트에 익숙함
- 인터넷이 불안정한 농촌지역에서도 모바일로 폼 제출이 가능
- 추후 데이터가 쌓이고 필요가 커지면 자체 웹앱(대시보드)으로 단계적으로 확장 가능

## 현재 단계 (MVP)

착수 100일 계획의 01~05번(거점 현황조사, baseline, 운영규칙, 진단도구, 출석양식)에
대응하는 데이터 구조를 먼저 만들었습니다. 실제 운영은 아래 순서로 진행합니다.

1. `docs/data-model.md` 구조로 Google Sheets 생성 (완료 — Drive에 업로드됨)
2. `docs/forms-spec.md` 명세대로 Google Forms 5종 생성 및 시트 연동
3. GYA Hub 스태프에게 시트 편집 권한 공유
4. 7개 공부방 실사 데이터 입력 → MVP 가동
5. 12주 파일럿 운영 후 08_대시보드 지표로 점검, 필요 시 자체 웹앱으로 확장 검토

## 안전 관련 안내 (Safeguarding)

아동 개인정보·사진은 이 공개 저장소(GitHub)에 올리지 않습니다. 학생 데이터는
비공개 Google Sheets에서만 관리하고, 사진 게재는 사진동의 절차(제안서 11장)를
확인한 뒤 별도 채널로만 공유합니다.
