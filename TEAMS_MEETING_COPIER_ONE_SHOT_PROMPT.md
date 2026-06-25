# Teams Meeting Copier 한 번에 완성 프롬프트

이 문서는 Teams 회의 정보 복사 앱을 처음부터 다시 만들 때 쓰는 전체 요구사항 프롬프트입니다. 특히 이번에 오래 막혔던 Microsoft 로그인, Teams 앱 내부 로그인, 기존 브라우저 로그인 세션, Netlify 캐시, Teams 앱 패키지 버전 문제를 처음부터 피하도록 정리했습니다.

## 최종 목표

Microsoft Teams 개인 앱으로 실행되는 정적 웹앱을 만든다.

사용자는 Microsoft 계정으로 로그인한 뒤 기간을 선택하고, Microsoft Graph Calendar API로 해당 기간의 Teams 회의를 조회한다. 조회 결과에서 회의 제목, 일시, Teams 링크를 선택 복사 또는 전체 복사할 수 있다.

배포는 Netlify 정적 사이트로 한다. Teams 앱은 Netlify URL을 개인 탭으로 여는 manifest 패키지를 사용한다.

## 완성 앱의 핵심 원칙

1. Teams 앱 안에서도 일반 웹과 같은 MSAL `loginPopup()` 흐름을 사용한다.
2. `microsoftTeams.authentication.authenticate()` 기반의 `auth-start.html` / `auth-end.html` Teams 전용 인증 흐름은 사용하지 않는다.
3. 앱 로드만으로 Microsoft 로그인 초기화를 실행하지 않는다.
4. 로그인은 사용자가 `Microsoft 로그인` 버튼을 누를 때만 시작한다.
5. 기존 브라우저에 남은 `interaction_in_progress` 상태를 정리한다.
6. MSAL 캐시는 `sessionStorage`를 사용하고 `storeAuthStateInCookie`는 `false`로 둔다.
7. 로그인 성공 후 계정과 access token을 앱 저장소에 저장해서 같은 브라우저/Teams 앱에서 다시 열어도 로그인 상태를 복원한다.
8. 토큰이 만료되면 다시 로그인 또는 토큰 팝업으로 갱신한다.
9. Netlify HTML/JS는 캐시 문제를 줄이기 위해 no-cache 헤더와 버전 쿼리를 사용한다.
10. Teams 앱 manifest 버전은 업로드할 때마다 반드시 증가시킨다.

## 한 번에 성공시키는 구현 프롬프트

아래 프롬프트를 그대로 새 작업에 사용한다.

```text
Netlify에 배포되는 정적 웹앱과 Microsoft Teams 개인 앱 패키지를 만들어줘.

앱 이름은 Teams Meeting Copier다.

목표:
- Microsoft 계정으로 로그인한다.
- Microsoft Graph /me/calendarView API로 지정 기간의 일정 중 Teams 회의를 조회한다.
- 회의 제목, 시작/종료 일시, Teams 링크를 화면에 표로 보여준다.
- 선택한 회의 또는 전체 회의를 클립보드에 복사한다.
- Teams 개인 앱 안에서도 동작해야 한다.
- 일반 브라우저에서도 동작해야 한다.

중요한 인증 요구:
- MSAL browser를 사용한다.
- Teams 앱 안에서도 `microsoftTeams.authentication.authenticate()`를 쓰지 않는다.
- Teams 앱 안에서도 일반 웹과 동일하게 `msalClient.loginPopup()`을 사용한다.
- `auth-start.html` / `auth-end.html` 기반 Teams 전용 인증 팝업 플로우를 만들지 않는다.
- 앱 로드 시 자동으로 로그인 redirect/silent 인증을 실행하지 않는다.
- 로그인은 사용자가 `Microsoft 로그인` 버튼을 클릭했을 때만 시작한다.
- `loginPopup()` 요청 scope는 `openid`, `profile`, `Calendars.Read`를 포함한다.
- 로그인 성공 시 `loginResult.account`를 activeAccount로 설정하고, `loginResult.accessToken`을 저장한다.
- 회의 조회 시에는 로그인 팝업에서 받은 access token을 먼저 사용한다.
- access token이 없거나 만료되었으면 `acquireTokenSilent()`를 시도한다.
- silent가 `interaction_required`, `login_required`, `consent_required` 등으로 실패하면 `acquireTokenPopup()`으로 토큰을 다시 받는다.
- `cacheLocation`은 `sessionStorage`로 설정한다.
- `storeAuthStateInCookie`는 `false`로 설정한다.
- 기존 브라우저에 남은 `interaction_in_progress` 문제를 피하기 위해 로그인 전 sessionStorage/localStorage의 MSAL interaction 상태와 MSAL 쿠키를 정리하는 함수를 둔다.
- 단, 정상 로그인 유지용 저장값은 앱 로드 때 지우지 않는다.

로그인 유지 요구:
- 로그인 성공 시 username, account id, access token, expiresOn을 localStorage에 저장한다.
- 앱 시작 시 저장된 토큰이 있고 만료되지 않았으면 로그인 상태를 즉시 복원한다.
- 복원되면 화면에 `로그인 계정: user@example.com`을 표시한다.
- 로그아웃 버튼을 누를 때만 저장된 로그인 상태를 삭제한다.
- 토큰 만료 1분 전이면 저장 로그인 상태를 버리고 다시 로그인하게 한다.

UI 요구:
- 첫 화면은 앱 실제 사용 화면이어야 한다. 랜딩 페이지를 만들지 않는다.
- 상단에 Teams 아이콘과 `TEAMS Meetings` 제목을 표시한다.
- 아이콘은 외부 SVG 파일 로딩에 의존하지 말고 inline SVG로 넣는다.
- 연결 설정 카드에는 `Microsoft 로그인`, `로그아웃`, 로그인 상태 문구를 둔다.
- 회의 조회 카드에는 시작 일시, 종료 일시 input과 `기간 조회`, `선택 복사`, `전체 복사` 버튼을 둔다.
- 조회 결과 표에는 선택 체크박스, 주제, 일시, 링크, 동작 열을 둔다.
- 복사 미리보기 textarea를 둔다.
- UI 문구는 한국어로 한다.

Graph 조회 요구:
- Graph endpoint:
  `https://graph.microsoft.com/v1.0/me/calendarView`
- query:
  - startDateTime: 시작 일시 ISO
  - endDateTime: 종료 일시 ISO
  - $orderby: start/dateTime
  - $top: 100
  - $select: subject,start,end,isOnlineMeeting,onlineMeeting,bodyPreview
- Teams 회의 판정:
  - `event.isOnlineMeeting`이 true거나
  - `event.onlineMeeting.joinUrl`이 있거나
  - `bodyPreview`에서 `https://teams.microsoft.com/...` 링크를 추출할 수 있으면 Teams 회의로 본다.
- 링크는 `event.onlineMeeting.joinUrl`을 우선 사용하고 없으면 bodyPreview에서 정규식으로 추출한다.

파일 구조:
- public/index.html
- public/styles.css
- public/app.js
- public/config.js
- public/privacy.html
- public/terms.html
- public/reset.html
- netlify.toml
- teams-app/manifest.json
- teams-app/color.png
- teams-app/outline.png
- package-teams-app.ps1

public/config.js:
```js
window.APP_CONFIG = {
  tenant: "organizations",
  clientId: "ENTRA_APPLICATION_CLIENT_ID"
};
```

MSAL 설정:
```js
msalClient = new msal.PublicClientApplication({
  auth: {
    clientId,
    authority: "https://login.microsoftonline.com/organizations",
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
});
```

절대 하지 말 것:
- 앱 로드 시 `loginRedirect()`를 자동 호출하지 않는다.
- 앱 로드 시 `handleRedirectPromise()`를 무조건 호출하지 않는다.
- Teams 내부 로그인용으로 `microsoftTeams.authentication.authenticate()`를 쓰지 않는다.
- Teams 팝업 전용 `/auth-start.html?teams=1` 흐름에 의존하지 않는다.
- `storeAuthStateInCookie: true`를 쓰지 않는다.
- 매 앱 로드마다 localStorage의 정상 로그인 저장값을 지우지 않는다.
- 제목 아이콘을 외부 SVG 파일에만 의존하지 않는다.

Netlify 설정:
- publish directory는 `public`
- build command는 비워둔다.
- HTML, app.js, config.js는 no-cache 헤더를 둔다.
- index.html에서 app.js/config.js 로드 시 버전 쿼리를 붙인다.

netlify.toml 예:
```toml
[build]
  publish = "public"

[[headers]]
  for = "/*.html"
  [headers.values]
    Cache-Control = "no-cache, no-store, must-revalidate"

[[headers]]
  for = "/app.js"
  [headers.values]
    Cache-Control = "no-cache, no-store, must-revalidate"

[[headers]]
  for = "/config.js"
  [headers.values]
    Cache-Control = "no-cache, no-store, must-revalidate"

[[redirects]]
  from = "/*"
  to = "/index.html"
  status = 200
```

Teams manifest 요구:
- manifestVersion: 1.16 이상
- staticTabs 개인 앱 personal scope
- contentUrl과 websiteUrl은 Netlify 루트 URL
- validDomains에는 Netlify 도메인을 넣는다.
- 필요하면 login.microsoftonline.com도 넣는다.
- 패키지를 다시 업로드할 때마다 `version`을 반드시 증가시킨다. 예: 1.0.2 -> 1.0.3

Entra 앱 등록 요구:
- Supported account types: 조직 계정 사용 요구에 맞게 선택한다. 여러 테넌트면 multi-tenant.
- Platform: Single-page application (SPA)
- Redirect URI:
  - `https://YOUR_NETLIFY_DOMAIN`
  - 로컬 테스트 시 `http://localhost:5500`
- API permission:
  - Microsoft Graph Delegated permission: `Calendars.Read`
- 필요한 테넌트마다 admin consent를 완료한다.

검증:
- 일반 브라우저에서 로그인 팝업이 뜨고 계정 표시가 되는지 확인한다.
- 기존에 M365에 로그인된 브라우저에서도 로그인 팝업이 정상 동작하는지 확인한다.
- Teams 웹 앱 내부에서도 로그인 팝업 후 Teams 앱 본문에 계정 표시가 되는지 확인한다.
- 조회 버튼 클릭 시 Graph 결과가 나오거나 권한/Graph 오류 메시지가 정확히 표시되는지 확인한다.
- 앱을 다른 화면으로 갔다가 다시 열어도 저장된 로그인 상태가 유지되는지 확인한다.
- 로그아웃 버튼을 누르면 저장 로그인 상태가 삭제되는지 확인한다.
```

## 반드시 피해야 할 로그인 함정

### 1. Teams SDK 인증 팝업에 의존하지 말 것

처음에는 Teams 앱 내부 로그인이라서 `microsoftTeams.authentication.authenticate()`를 사용하려고 했다. 그러나 Teams 웹/브라우저 환경에서는 팝업이 `auth-start.html`이 아니라 루트 앱 화면으로 떨어지거나, `auth-start -> auth-end` 사이에서 팝업 상태가 사라지는 문제가 반복되었다.

최종적으로 성공한 방식은 Teams 앱 안에서도 일반 웹과 동일하게 MSAL `loginPopup()`을 직접 사용하는 것이었다.

### 2. 앱 로드 시 인증을 자동 실행하지 말 것

기존 M365 로그인 세션이 있는 브라우저에서 앱 로드 시 `handleRedirectPromise()`나 silent 인증을 무조건 실행하면 `interaction_in_progress` 루프가 발생할 수 있다.

앱 로드 시에는 화면을 먼저 띄우고, 저장된 앱 로그인 상태만 복원한다. 새 Microsoft 로그인은 버튼 클릭 때만 시작한다.

### 3. `storeAuthStateInCookie: true`를 쓰지 말 것

이 옵션이 켜져 있으면 MSAL interaction 상태가 쿠키에도 남아 기존 브라우저에서 계속 `interaction_in_progress`가 되살아날 수 있다.

최종 설정은 다음과 같다.

```js
cache: {
  cacheLocation: "sessionStorage",
  storeAuthStateInCookie: false
}
```

### 4. 정상 로그인 저장값과 문제 interaction 플래그를 구분할 것

로그인 문제를 해결한다고 localStorage/sessionStorage를 매번 전체 삭제하면 앱을 다시 열 때마다 로그인이 풀린다.

정리 대상:
- MSAL interaction status
- `interaction_in_progress`
- 필요 시 MSAL 임시 캐시

유지 대상:
- 앱이 직접 저장한 username
- 앱이 직접 저장한 아직 유효한 access token
- token expiresOn

### 5. 로그인 팝업 결과의 access token을 버리지 말 것

로그인 팝업에서 `Calendars.Read` scope를 포함해 성공하면 `loginResult.accessToken`이 온다. 이 토큰을 버리고 조회 때 다시 `acquireTokenSilent()`만 시도하면 권한/세션 오류가 날 수 있다.

최종 흐름:

1. `loginPopup(["openid", "profile", "Calendars.Read"])`
2. `loginResult.account` 저장
3. `loginResult.accessToken` 저장
4. 조회 시 저장 token 우선 사용
5. 만료/부족하면 `acquireTokenSilent()`
6. silent 실패 시 `acquireTokenPopup()`

### 6. 제목 아이콘은 inline SVG로 넣을 것

Teams/팝업 렌더링에서 외부 SVG 파일이 빠질 수 있다. 제목 옆 아이콘은 inline SVG로 넣으면 가장 안정적이다.

## Microsoft Entra 체크리스트

Entra 앱 등록에서 아래를 확인한다.

- Application client ID가 `public/config.js`에 들어갔는가
- SPA Redirect URI에 Netlify 루트가 들어갔는가
- 로컬 테스트를 한다면 `http://localhost:5500`도 들어갔는가
- Microsoft Graph delegated `Calendars.Read` 권한이 있는가
- 필요한 테넌트에서 admin consent가 완료되었는가
- 사용자가 해당 Enterprise Application에 접근 가능한가
- User assignment required가 켜져 있다면 사용자/그룹이 할당되었는가

## Teams 앱 패키징 체크리스트

`teams-app/manifest.json`에서 확인한다.

```json
{
  "manifestVersion": "1.16",
  "version": "1.0.3",
  "staticTabs": [
    {
      "contentUrl": "https://teamsmeetings.netlify.app",
      "websiteUrl": "https://teamsmeetings.netlify.app",
      "scopes": ["personal"]
    }
  ],
  "validDomains": ["teamsmeetings.netlify.app", "login.microsoftonline.com"]
}
```

패키징:

```powershell
powershell -ExecutionPolicy Bypass -File .\package-teams-app.ps1
```

업로드 시 주의:

- Teams 관리 센터의 기존 게시 버전보다 높은 manifest `version`이어야 한다.
- 같은 버전 zip을 업데이트로 올리면 거절된다.
- 예: 게시 버전 1.0.2면 새 파일은 1.0.3 이상이어야 한다.
- 웹 코드만 바뀐 경우에는 Teams 패키지 재업로드가 필요 없다.
- manifest, 앱 이름, 아이콘, 도메인, URL이 바뀌면 패키지를 새로 만들어 업로드한다.

## Netlify 체크리스트

- GitHub main 브랜치와 연결되어 있는가
- Publish directory가 `public`인가
- 최신 커밋이 Published 되었는가
- 브라우저가 오래된 JS를 잡지 않도록 app.js에 버전 쿼리가 붙어 있는가
- `netlify.toml`에 no-cache 헤더가 있는가

## 문제별 빠른 진단

### 루트 주소에서 빈 화면과 앱 화면이 반복된다

원인:
- 앱 로드 시 자동 인증 처리 또는 오래된 JS 캐시

해결:
- 앱 로드 시 `loginRedirect`, `handleRedirectPromise`, silent 인증을 무조건 실행하지 않는다.
- app.js 버전 쿼리를 올린다.
- no-cache 헤더를 설정한다.

### `interaction_in_progress`가 계속 뜬다

원인:
- MSAL interaction 상태가 sessionStorage/localStorage/cookie에 남음

해결:
- interaction 상태만 지우는 함수를 둔다.
- `storeAuthStateInCookie: false`
- MSAL cookie 중 `msal` 포함 이름은 reset 시 삭제한다.

### 시크릿창은 되는데 기존 브라우저만 안 된다

원인:
- 기존 M365 SSO 쿠키 또는 기존 MSAL 캐시 충돌

해결:
- 로그인 전 MSAL 캐시와 interaction 상태를 정리한다.
- 팝업 로그인 사용
- 필요 시 `/reset.html`로 앱 저장소 초기화

### Teams 앱 안에서 팝업이 앱 화면만 보여준다

원인:
- Teams SDK 인증 팝업 흐름이 `auth-start.html`이 아니라 루트 화면으로 빠짐
- 또는 `auth-start -> auth-end` 사이에서 Teams 팝업 플래그가 사라짐

권장 해결:
- Teams SDK 인증 팝업 흐름을 쓰지 않는다.
- Teams 안에서도 `msalClient.loginPopup()`을 직접 사용한다.

### 로그인은 됐는데 회의 조회에서 세션/권한 오류가 난다

원인:
- 로그인 팝업에서 받은 access token을 저장하지 않고, 조회 때 silent token만 시도함

해결:
- `loginResult.accessToken`을 저장한다.
- 조회 때 저장 token을 먼저 사용한다.
- 만료 시 silent, 필요 시 popup으로 갱신한다.

### Teams 관리 센터에서 버전 오류가 난다

원인:
- 이미 게시된 manifest version과 같은 버전의 zip을 업로드함

해결:
- `teams-app/manifest.json`의 `version`을 증가시킨다.
- 패키지를 다시 만든다.

## 현재 성공 기준

성공한 앱은 아래 조건을 만족해야 한다.

- 일반 브라우저에서 로그인 성공
- 기존 M365 로그인 브라우저에서도 로그인 성공
- Teams 웹 앱 내부에서도 로그인 성공
- 로그인 후 Teams 앱 본문에 계정 표시
- 기간 조회 시 Graph 회의 목록 표시
- 앱을 닫았다 다시 열어도 토큰이 유효하면 로그인 유지
- 로그아웃을 눌렀을 때만 저장 로그인 상태 삭제
- Teams 패키지 버전 업데이트 가능

