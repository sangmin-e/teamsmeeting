# Re-Register From Scratch (Recommended)

Use this when admin approval keeps repeating.

## 1) App registration reset

1. In the home tenant portal, open App registrations.
2. Delete old TEAMS Meeting app (or create a new app with a new name).
3. Create a new app registration with:
- Supported account types: Accounts in any organizational directory
- Platform: Single-page application (SPA)
- Redirect URI (local): http://localhost:5500
- Redirect URI (Teams auth): https://YOUR_NETLIFY_DOMAIN/auth-end.html

## 2) API permissions

1. Add Microsoft Graph delegated permission: Calendars.Read
2. Save.

## 3) Copy new Client ID

1. From Overview, copy Application (client) ID.
2. Update public/config.js clientId value with the new ID.

## 4) Tenant admin consent (must be per tenant)

For each tenant admin account, open admin-consent URL:
- https://login.microsoftonline.com/ogeum.sen.ms.kr/adminconsent?client_id=NEW_CLIENT_ID&redirect_uri=http%3A%2F%2Flocalhost%3A5500
- https://login.microsoftonline.com/microedu.kr/adminconsent?client_id=NEW_CLIENT_ID&redirect_uri=http%3A%2F%2Flocalhost%3A5500
- https://login.microsoftonline.com/copilots.kr/adminconsent?client_id=NEW_CLIENT_ID&redirect_uri=http%3A%2F%2Flocalhost%3A5500

## 5) Verify in each tenant portal

In each tenant:
1. Enterprise applications > All applications > TEAMS Meeting exists.
2. Application ID equals your NEW_CLIENT_ID.
3. Permissions shows Graph Calendars.Read as Granted.
4. If User assignment required is Yes, assign users/groups or set No.

## 6) Test

1. Open app in an incognito window.
2. Sign in with one tenant account at a time.
3. Repeat for all target tenants.

## 7) Netlify after local success

1. Add SPA redirect URI: https://YOUR_NETLIFY_DOMAIN
2. Add SPA redirect URI: https://YOUR_NETLIFY_DOMAIN/auth-end.html
3. Repeat admin consent per tenant using redirect_uri=https://YOUR_NETLIFY_DOMAIN
