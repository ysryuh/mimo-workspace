# MEMORY.md - 장기 기억

## 프로젝트: 고하중 AMMR 개발 (AMMR-H600)
- 양팔 + AMR 적재 ≥600kg + 팔 가반 ≥30kg + 자율주행
- "양팔 + 30kg+ + 600kg AMR" 통합 제품 시장에 없음 → 블루오션
- ROKAE가 가장 가까운 경쟁자 (CMR 500kg, CR35 45kg 가반)
- AMR 600kg 이상 자체 개발 필요 (ROKAE CMR은 500kg 최대)
- **제품명: AMMR-H600**

### 핵심 기술 결정사항
- **구동 방식**: Steering drive (Swerve drive) — 메카넘 휠 대신 채택
- **바퀴**: 고탄성 고무바퀴, Φ200mm (8인치), 폭 50~65mm
  - 메카넘 휠은 바닥 상태 영향 + 진동 문제로 제외
- **바퀴 계산**: 급정지 1m 기준 제동토크 28.1 N·m, 구동출력 ~400W/바퀴
- **배터리 충전**: LiFePO4, 급속충전 "80% in 45분"으로 조정 (30분은 특수 셀 필요)
- **적재 하중 분산**: 바퀴당 250kg

### 개발 로드맵
- **6개월 프로토타입** (기존 24개월에서 압축)
- 자체 제어 시스템 & 자율주행 플랫폼 이미 보유
- 양팔은 COTS (기성품) 선택
- HW/SW 병렬 진행
  - W1~4: 설계 확정 + 부품 소싱
  - W2~8: HW 통합
  - W4~16: SW 통합
  - W12~22: 테스트 & 최적화
  - W20~26: 데모 & 파일럿

### 작업 산출물
- `AMMR_통합설계서.md` — 통합설계서 v0.2 (마크다운)
- `AMMR_H600_통합설계서.pptx` — PPT 버전 (15 슬라이드)
- `create_ppt.py` — python-pptx 기반 PPT 생성 스크립트
- Gofile 링크: https://gofile.io/d/2Zt2nn

## 제임스 (James)
- 한국어 소통, 타임존 Asia/Shanghai (GMT+8)
- GitHub: ysryuh/mimo-workspace (private)
- 직접 수정 선호 — "내가 수정할께" 스타일
- 기술적 검증 요청 잘함 (배터리 충전율 질문 등)

## 세션 이력
- 2026-03-31: 첫 세션, BOOTSTRAP.md 삭제, IDENTITY/USER/SOUL 설정, ROKAE 조사
- 2026-04-01: 통합설계서 작성, PPT 생성, 로드맵 6개월로 수정, 구동시스템/바퀴 계산, GitHub push 설정
