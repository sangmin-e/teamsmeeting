const startDateEl = document.getElementById("startDate");
const endDateEl = document.getElementById("endDate");
const statusEl = document.getElementById("status");
const accountInfoEl = document.getElementById("accountInfo");
const meetingRowsEl = document.getElementById("meetingRows");
const emptyStateEl = document.getElementById("emptyState");
const previewTextEl = document.getElementById("previewText");

const signInBtn = document.getElementById("signInBtn");
const signOutBtn = document.getElementById("signOutBtn");
const fetchBtn = document.getElementById("fetchBtn");
const copySelectedBtn = document.getElementById("copySelectedBtn");
const copyAllBtn = document.getElementById("copyAllBtn");

const GRAPH_SCOPES = ["Calendars.Read"];
const REDIRECT_FLAG_KEY = "teamsMeeting.authRedirectInFlight";
const LAST_ACCOUNT_KEY = "teamsMeeting.lastAccount";
const MSAL_CDN_URLS = [
  "https://alcdn.msauth.net/browser/2.38.4/js/msal-browser.min.js",
  "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.4/lib/msal-browser.min.js",
  "https://unpkg.com/@azure/msal-browser@2.38.4/lib/msal-browser.min.js",
];

let msalClient = null;
let activeAccount = null;
let meetings = [];
let msalLoaderPromise = null;
let authReadyPromise = null;

function getAppConfig() {
  const config = window.APP_CONFIG || {};
  const tenant = (config.tenant || "common").trim();
  const clientId = (config.clientId || "").trim();
  return { tenant, clientId };
}

function hasValidClientId(clientId) {
  if (!clientId) {
    return false;
  }

  if (clientId === "REPLACE_WITH_ENTRA_CLIENT_ID") {
    return false;
  }

  // Basic GUID pattern used by Entra Application (client) ID.
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    clientId,
  );
}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle("error", isError);
}

function rememberAccount(account) {
  if (!account?.username) {
    return;
  }

  localStorage.setItem(LAST_ACCOUNT_KEY, account.username);
}

function pickPreferredAccount(accounts) {
  if (!accounts?.length) {
    return null;
  }

  const last = localStorage.getItem(LAST_ACCOUNT_KEY);
  if (!last) {
    return accounts[0];
  }

  return accounts.find((item) => item.username === last) || accounts[0];
}

function toUserErrorMessage(error, fallback) {
  const code = error?.errorCode || error?.code || "";
  const raw = String(error?.message || "");

  if (code === "interaction_in_progress" || raw.includes("interaction_in_progress")) {
    return "로그인 진행 중입니다. 열린 Microsoft 로그인 창을 먼저 완료한 뒤 다시 시도하세요.";
  }

  if (
    code === "interaction_required" ||
    code === "consent_required" ||
    code === "login_required" ||
    raw.includes("interaction_required")
  ) {
    return "세션이 만료되었거나 권한 재확인이 필요합니다. Microsoft 로그인을 다시 실행하세요.";
  }

  if (raw.includes("Need admin approval") || raw.includes("AADSTS65001")) {
    return "관리자 동의가 아직 반영되지 않았습니다. 아래 관리자 동의 링크로 테넌트별 승인 후 다시 로그인하세요.";
  }

  return raw || fallback;
}

function buildAuthority() {
  const tenantInput = getAppConfig().tenant;
  if (!tenantInput) {
    return "https://login.microsoftonline.com/common";
  }

  if (tenantInput.includes("@")) {
    return "https://login.microsoftonline.com/organizations";
  }

  return `https://login.microsoftonline.com/${tenantInput}`;
}

function setDefaultDateRange() {
  const now = new Date();
  const end = new Date(now);
  end.setDate(now.getDate() + 7);

  startDateEl.value = toDateTimeLocal(now);
  endDateEl.value = toDateTimeLocal(end);
}

function toDateTimeLocal(date) {
  const pad = (value) => String(value).padStart(2, "0");

  const yyyy = date.getFullYear();
  const mm = pad(date.getMonth() + 1);
  const dd = pad(date.getDate());
  const hh = pad(date.getHours());
  const min = pad(date.getMinutes());

  return `${yyyy}-${mm}-${dd}T${hh}:${min}`;
}

function toDisplayDateTime(value) {
  if (!value?.dateTime) {
    return "(일시 없음)";
  }

  const dt = new Date(value.dateTime);
  return dt.toLocaleString("ko-KR", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function extractTeamsUrl(text) {
  if (!text) {
    return "";
  }

  const match = text.match(/https:\/\/teams\.microsoft\.com\/[^\s"')<]+/i);
  return match ? match[0] : "";
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`failed: ${src}`));
    document.head.appendChild(script);
  });
}

async function ensureMsalLoaded() {
  if (window.msal) {
    return;
  }

  if (msalLoaderPromise) {
    return msalLoaderPromise;
  }

  msalLoaderPromise = (async () => {
    for (const url of MSAL_CDN_URLS) {
      try {
        await loadScript(url);
        if (window.msal) {
          return;
        }
      } catch {
        // Try the next CDN.
      }
    }

    throw new Error("로그인 모듈을 불러오지 못했습니다. 잠시 후 다시 시도하세요.");
  })();

  return msalLoaderPromise;
}

async function createMsalClient() {
  const clientId = getAppConfig().clientId;
  if (!hasValidClientId(clientId)) {
    throw new Error("로그인 설정이 완료되지 않았습니다. 관리자에게 문의하세요.");
  }

  await ensureMsalLoaded();

  if (!["http:", "https:"].includes(window.location.protocol)) {
    throw new Error("로그인은 웹 주소(https 또는 localhost)에서만 사용할 수 있습니다.");
  }

  msalClient = new window.msal.PublicClientApplication({
    auth: {
      clientId,
      authority: buildAuthority(),
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true,
    },
  });

  const accounts = msalClient.getAllAccounts();
  activeAccount = pickPreferredAccount(accounts);
  if (activeAccount) {
    msalClient.setActiveAccount(activeAccount);
    rememberAccount(activeAccount);
  }
  renderAccountInfo();
}

async function ensureAuthReady() {
  if (authReadyPromise) {
    return authReadyPromise;
  }

  authReadyPromise = (async () => {
    await createMsalClient();

    const redirectResult = await msalClient.handleRedirectPromise();
    sessionStorage.removeItem(REDIRECT_FLAG_KEY);
    if (redirectResult?.account) {
      activeAccount = redirectResult.account;
      msalClient.setActiveAccount(activeAccount);
      rememberAccount(activeAccount);
    }

    if (!activeAccount) {
      const accounts = msalClient.getAllAccounts();
      activeAccount = pickPreferredAccount(accounts);
      if (activeAccount) {
        msalClient.setActiveAccount(activeAccount);
        rememberAccount(activeAccount);
      }
    }

    renderAccountInfo();
  })();

  return authReadyPromise;
}

function renderAccountInfo() {
  if (!activeAccount) {
    accountInfoEl.textContent = "로그인되지 않았습니다.";
    return;
  }

  accountInfoEl.textContent = `로그인 계정: ${activeAccount.username}`;
}

async function signIn() {
  await ensureAuthReady();

  if (activeAccount) {
    setStatus("이미 로그인되어 있습니다. 기간 조회를 실행하세요.");
    return;
  }

  if (sessionStorage.getItem(REDIRECT_FLAG_KEY) === "1") {
    setStatus("로그인 진행 중입니다. 열린 Microsoft 로그인 창을 완료하세요.", true);
    return;
  }

  sessionStorage.setItem(REDIRECT_FLAG_KEY, "1");

  await msalClient.loginRedirect({
    scopes: ["openid", "profile", ...GRAPH_SCOPES],
    prompt: "select_account",
    redirectStartPage: window.location.href,
  });
}

async function signOut() {
  await ensureAuthReady();

  if (activeAccount) {
    await msalClient.logoutRedirect({
      account: activeAccount,
      postLogoutRedirectUri: window.location.origin,
    });
    return;
  }

  activeAccount = null;
  localStorage.removeItem(LAST_ACCOUNT_KEY);
  renderAccountInfo();
  setStatus("로그아웃되었습니다.");
}

async function getAccessToken() {
  await ensureAuthReady();

  if (!msalClient || !activeAccount) {
    throw new Error("먼저 로그인하세요.");
  }

  try {
    const silent = await msalClient.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: activeAccount,
    });

    rememberAccount(activeAccount);

    return silent.accessToken;
  } catch (error) {
    throw new Error(toUserErrorMessage(error, "토큰을 가져오지 못했습니다. Microsoft 로그인을 다시 시도하세요."));
  }
}

async function fetchMeetings() {
  const startValue = startDateEl.value;
  const endValue = endDateEl.value;

  if (!startValue || !endValue) {
    throw new Error("시작/종료 일시를 입력하세요.");
  }

  const start = new Date(startValue);
  const end = new Date(endValue);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    throw new Error("일시 형식이 올바르지 않습니다.");
  }

  if (end <= start) {
    throw new Error("종료 일시는 시작 일시보다 늦어야 합니다.");
  }

  const token = await getAccessToken();
  const query = new URLSearchParams({
    startDateTime: start.toISOString(),
    endDateTime: end.toISOString(),
    "$orderby": "start/dateTime",
    "$top": "100",
    "$select": "subject,start,end,isOnlineMeeting,onlineMeeting,bodyPreview",
  });

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/calendarView?${query.toString()}`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    },
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph 조회 실패: ${response.status} ${text}`);
  }

  const data = await response.json();
  const rawEvents = Array.isArray(data.value) ? data.value : [];

  meetings = rawEvents
    .map((event) => {
      const joinUrl = event.onlineMeeting?.joinUrl || extractTeamsUrl(event.bodyPreview);
      const isTeamsMeeting =
        event.isOnlineMeeting || (joinUrl && joinUrl.toLowerCase().includes("teams.microsoft.com"));

      if (!isTeamsMeeting || !joinUrl) {
        return null;
      }

      return {
        id: event.id,
        subject: event.subject || "(제목 없음)",
        start: toDisplayDateTime(event.start),
        end: toDisplayDateTime(event.end),
        link: joinUrl,
      };
    })
    .filter(Boolean);

  renderMeetings();
  setStatus(`${meetings.length}개의 Teams 회의를 조회했습니다.`);
}

function renderMeetings() {
  meetingRowsEl.innerHTML = "";
  emptyStateEl.style.display = meetings.length ? "none" : "block";

  meetings.forEach((meeting) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><input type="checkbox" class="meeting-check" data-id="${meeting.id}" /></td>
      <td>${meeting.subject}</td>
      <td>${meeting.start} ~ ${meeting.end}</td>
      <td><a href="${meeting.link}" target="_blank" rel="noreferrer">링크 열기</a></td>
      <td><button class="btn small secondary copy-one" data-id="${meeting.id}">이 회의 복사</button></td>
    `;
    meetingRowsEl.appendChild(row);
  });
}

function formatMeeting(meeting, index) {
  return [
    `- 회의 주제: ${meeting.subject}`,
    `- 회의 일시: ${meeting.start} ~ ${meeting.end}`,
    `- 회의 링크: ${meeting.link}`,
  ].join("\n");
}

function selectedMeetings() {
  const selectedIds = new Set(
    Array.from(document.querySelectorAll(".meeting-check:checked")).map((el) => el.dataset.id),
  );

  return meetings.filter((meeting) => selectedIds.has(meeting.id));
}

async function copyMeetings(targetMeetings) {
  if (!targetMeetings.length) {
    throw new Error("복사할 회의를 선택하세요.");
  }

  const text = targetMeetings.map((meeting, index) => formatMeeting(meeting, index)).join("\n\n");
  previewTextEl.value = text;
  await navigator.clipboard.writeText(text);
}

signInBtn.addEventListener("click", async () => {
  try {
    await signIn();
  } catch (error) {
    setStatus(toUserErrorMessage(error, "로그인 실패"), true);
  }
});

signOutBtn.addEventListener("click", async () => {
  try {
    await signOut();
  } catch (error) {
    setStatus(toUserErrorMessage(error, "로그아웃 실패"), true);
  }
});

fetchBtn.addEventListener("click", async () => {
  try {
    await fetchMeetings();
  } catch (error) {
    setStatus(toUserErrorMessage(error, "조회 실패"), true);
  }
});

copySelectedBtn.addEventListener("click", async () => {
  try {
    await copyMeetings(selectedMeetings());
    setStatus("선택한 회의 정보를 클립보드에 복사했습니다.");
  } catch (error) {
    setStatus(error.message || "선택 복사 실패", true);
  }
});

copyAllBtn.addEventListener("click", async () => {
  try {
    await copyMeetings(meetings);
    setStatus("조회된 모든 회의 정보를 클립보드에 복사했습니다.");
  } catch (error) {
    setStatus(error.message || "전체 복사 실패", true);
  }
});

meetingRowsEl.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement) || !target.classList.contains("copy-one")) {
    return;
  }

  const meetingId = target.dataset.id;
  const meeting = meetings.find((item) => item.id === meetingId);
  if (!meeting) {
    return;
  }

  try {
    await copyMeetings([meeting]);
    setStatus("선택한 회의 1건을 복사했습니다.");
  } catch (error) {
    setStatus(error.message || "개별 복사 실패", true);
  }
});

async function initApp() {
  setDefaultDateRange();
  renderAccountInfo();

  try {
    await ensureAuthReady();
    if (activeAccount) {
      setStatus("로그인 상태입니다. 기간 조회를 실행하세요.");
    } else {
      setStatus("Microsoft 로그인 후 기간 조회를 실행하세요.");
    }
  } catch (error) {
    setStatus(error.message || "로그인 초기화 실패", true);
  }
}

initApp();
