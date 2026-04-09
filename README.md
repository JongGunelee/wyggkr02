# wyggkr02

웹 대시보드와 로컬 브리지(EXE) 기반 자동화 실행 저장소입니다.

## 구성
- `00 dashboard.html`: GitHub Pages 웹 대시보드 진입점
- `automated_scripts/`: 카드 섹션에서 사용하는 자동화 스크립트
- `automated_app/`: 스크립트 실행에 필요한 앱/의존 리소스
- `dev_source/`: 빌드/배포 스크립트, 인수인계 문서, 런타임 저장소
- `runtime_store/`: 로컬 런타임 캐시/보조 패키지 저장 폴더
- `system_guides/`: 운영 및 구현 가이드 문서

## 배포 원칙
- GitHub 저장소를 기준 소스로 사용합니다.
- 설치형 ZIP은 `dev_source/runtime_store/WYGGKR02_Dashboard_Agent_Setup.zip` 경로를 기본 배포 경로로 사용합니다.
