# Teams Meeting Copier

MS Teams 일정에서 기간별 회의를 조회하고, 회의 주제/일시/링크를 선택 복사하는 Netlify 기반 정적 웹앱입니다.

사용자 화면은 로그인 버튼만 보이는 방식으로 동작합니다.

## 권장 진행 순서

1. Netlify 배포를 먼저 완료
2. 확정된 Netlify 도메인으로 Entra Redirect URI 등록
3. `public/config.js`의 clientId 반영
4. 같은 도메인으로 Teams manifest 값 반영 후 패키징

상세 순서: `DEPLOY_ORDER.md`

## 1) 로컬 확인

MS 로그인/Graph 조회를 위해서는 file:// 직접 실행이 아니라 localhost 실행이 필요합니다.

```powershell
./start-local-server.ps1
```

브라우저에서 `http://localhost:5500` 접속

로그인/조회 기능을 쓰려면 먼저 Microsoft Entra 앱 등록이 필요합니다.

## 2) Microsoft Entra 앱 등록

1. Azure Portal > Microsoft Entra ID > App registrations > New registration
2. Supported account types 선택
3. Redirect URI (SPA) 추가
4. 로컬 테스트: `http://localhost:5500` 또는 사용 중인 로컬 주소
5. 배포 주소: `https://YOUR_NETLIFY_DOMAIN`
6. Teams 로그인 완료 주소: `https://YOUR_NETLIFY_DOMAIN/auth-end.html`
7. API permissions > Microsoft Graph > Delegated: `Calendars.Read`
8. 필요 시 Admin consent 수행

앱에서 사용할 값:

- `Application (client) ID`

설정 위치(관리자 1회 설정):

- `public/config.js`의 `tenant`, `clientId`

## 3) Netlify 배포

1. 이 저장소를 Netlify에 연결합니다.
2. Build command: 비움 (정적 사이트)
3. Publish directory: `public`
4. 배포 후 발급된 도메인을 기록합니다. (예: `my-team-meeting.netlify.app`)

## 4) Teams 앱 매니페스트 수정

`teams-app/manifest.json`에서 아래 항목을 실제 값으로 바꿉니다.

- `id`: 새 GUID
- `websiteUrl`, `privacyUrl`, `termsOfUseUrl`, `contentUrl`
- `validDomains`: Netlify 도메인만 입력 (스킴 없이)

## 5) 아이콘 생성 및 패키지 만들기

`teams-app` 폴더에는 다음 파일이 필요합니다.

- `color.png` (192x192)
- `outline.png` (32x32)

패키지 생성:

```powershell
./package-teams-app.ps1
```

생성 결과:

- `teams-app-package.zip`

## 6) Teams에 업로드

1. Teams 관리 센터 또는 Teams 클라이언트에서 사용자 지정 앱 업로드를 허용합니다.
2. `teams-app-package.zip` 업로드
3. 개인 앱으로 설치 후 탭 실행

## 사용 흐름

1. 관리자: `public/config.js`에 clientId 설정
2. 사용자: Microsoft 로그인
3. 시작/종료 일시 입력 후 기간 조회
4. 조회된 회의 중 선택 복사 또는 전체 복사

## 배포 모드 원칙

1. 최종 사용자에게 Tenant/Client ID 입력을 받지 않습니다.
2. 관리자만 배포 전에 `public/config.js`를 1회 설정합니다.
3. Teams 탭에서는 로그인 버튼만 누르면 조회가 가능해야 합니다.
