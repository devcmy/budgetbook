# 테트리스 가계부 · iPhone 17 Pro Max 디자인 02

기존 엑셀 가계부의 기준을 유지하면서 개인용 iOS 화면을 다듬었습니다. 이번달은 **비고정비·고정비·비예비비·예비비 사용**을 먼저 보여주고, 설정 탭에서 **5개 탭과 25개 메뉴/표시 항목의 이름**을 바꿉니다. 앱 이름은 기존 엑셀 파일명에서 가져온 작업명입니다.

이 문서의 경로와 명령은 `app/` 폴더 기준입니다. 저장소 루트에서 `cd app`을 실행한 뒤 따라가세요.

## 바로 보기

- `design/preview.html`: 인터넷·별도 설치 없이 브라우저에서 여는 단일 파일 시안.
- `design/screenshots/04-design-board-v2.png`: 이번달·비예비비·개인 설정을 모은 이미지.
- `design/board.html`: 화면별 시안 보드.
- `design/디자인_명세.md`: 분류·설정·Glass 적용 근거.
- `design/apple-references.json`: 확인한 Apple 공식 가이드.

시안은 iPhone 17 Pro Max의 440×956pt를 기준으로 제작했습니다. 브라우저의 blur·투명 효과는 네이티브 Liquid Glass를 설명하기 위한 표현입니다.

## iOS 실행

Mac에서 `ios/TetrisBudget.xcodeproj`를 엽니다. **Glass API의 컴파일은 Xcode 26 이상 SDK가 필요**하며, iOS 27 기기·시뮬레이터 확인은 iOS 27을 지원하는 최신 Xcode/SDK에서 진행합니다. `TetrisBudget` 공유 스킴과 iPhone 17 Pro Max를 선택합니다. 실기기는 Signing & Capabilities의 Team/Bundle Identifier를 본인 환경에 맞게 설정합니다.

SwiftUI 시스템 `TabView`를 사용해 최신 iOS의 탭 디자인을 따릅니다. 월 선택·입력 버튼은 `glassEffect(.regular.interactive(), in: Capsule())`를 사용하고, 접근성의 투명도 감소와 앱의 투명 효과 감소를 반영합니다. 타깃 최소 버전은 기존 iOS 17을 유지했으며 iOS 26 미만에서는 표준 material로 표시합니다. 실제 앱 아이콘·배포 서명·TestFlight/App Store 설정은 아직 없습니다.

**Linux 환경이므로 Xcode 빌드, UserDefaults의 실기기 저장, 네이티브 Glass/키보드/안전영역은 검증하지 못했습니다.** 브라우저 시안의 상호작용·이름 변경·화면 크기를 검증했으며, 최종 iPhone 렌더링은 Xcode에서 확인해야 합니다.

## 현재 범위

5개 탭 이동, 이번달 네 지출 메뉴 상세, 일별 금액과 거래 상세, 검색·달력, 예비비·WRAP 잔액, 테트리스 비교를 구현했습니다. 설정은 탭 이름(12자), 메뉴/표시 이름(20자), 개별/전체 이름 복원, 촘촘한 보기, 투명 효과 감소를 제공합니다. 이름은 화면 표시만 바꾸며 원본 계좌·분류·거래 ID·금액은 유지합니다.

설정은 iOS의 UserDefaults, 브라우저의 localStorage에 저장하도록 구현했습니다. 브라우저 저장소가 제한되면 현재 화면에는 적용하고 저장 불가를 알립니다. 반면 거래 시안 입력은 계속 메모리에만 유지되며 재실행/새로고침 시 초기화됩니다. 원본 스냅샷은 읽기 전용이고 새 입력이 실제 결산·예비비·잔액을 재계산하지 않습니다.

다음 개발은 영구 거래 저장, Excel/RAW 이관, 중복을 막는 회계 계산 모듈, 정산·목적자금 연결 및 백업/복원입니다. 이번 변경은 전체 가계부 앱의 완성이나 회계 승인 사례 통과를 의미하지 않습니다.

## 검증

Python Playwright/Chromium 환경에서 `python design/verify_preview.py`를 실행합니다. 로컬 HTML을 메모리에서 렌더링하여 서버나 외부 네트워크 요청 없이 확인합니다. 결과는 `design/verification-v2.json`입니다.

이름 저장·복원 테스트는 검증환경의 저장소 제한 때문에 Web Storage API 대역을 사용했습니다. 실제 브라우저 디스크 지속성이나 iOS UserDefaults 실행 검증으로 해석하면 안 됩니다. Pro Max 440×956 및 390/375/320px의 5탭·상세 화면 넘침을 확인합니다.

원본 Excel과 인계 파일은 수정하지 않았고, 외부 공개·호스팅·배포를 하지 않았습니다. 01 버전 다운로드 ZIP도 유지했습니다.
