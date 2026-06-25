# Netlify First Deployment Order

## Recommended sequence

1. Deploy this project to Netlify first.
2. Confirm your final Netlify domain (for example: your-app.netlify.app).
3. Register/update Microsoft Entra app with that exact domain in Redirect URI (SPA).
4. Put real values into public/config.js.
5. Update teams-app/manifest.json with the same Netlify domain.
6. Rebuild Teams app package and upload to Teams.

## 1) Netlify deploy

- Publish directory: public
- Build command: (empty)

## 2) Entra app settings

- Authentication > Single-page application Redirect URI
  - https://YOUR_NETLIFY_DOMAIN
  - https://YOUR_NETLIFY_DOMAIN/auth-end.html
- API permissions
  - Microsoft Graph Delegated: Calendars.Read
- Copy Application (client) ID

## 3) App config

Edit public/config.js

- tenant: common (or your tenant)
- clientId: real Application (client) ID

## 4) Teams manifest settings

Edit teams-app/manifest.json

- developer.websiteUrl
- developer.privacyUrl
- developer.termsOfUseUrl
- staticTabs[0].contentUrl
- staticTabs[0].websiteUrl
- validDomains[0]

All above values must use the same Netlify domain.

## 5) Package and upload

Run:

./package-teams-app.ps1

Output:

- teams-app-package.zip

Upload teams-app-package.zip to Teams custom apps.
