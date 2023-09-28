# Sharepoint React App for SP Online / MS 365 and SP OnPrem and Subscription Edition (SE)

Example SP react app that reads list and adds items to list. The app's two versions:

1. SP OnLine SPFx react webpart

<p>Webpart uses @microsoft/sharepoint-generator but seperates the react app which provides ability to choose react versions different than dependency in @microsoft/sharepoint-generator.  Requires [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) in order to access SP Online site's information, specifically list data in this example app.  The MS Graph API be be be granted **Sites.ReadWrite.All** permissions, see [MS SP Admin](https://learn.microsoft.com/en-us/sharepoint/sharepoint-admin-role) for details.  Link to development tenant (access is limited but relative path will be same): Development SP Tenant's [SP Admin API](https://8r1bcm-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement) access.</p>

TODO Image here for webApiPermissionManagement

2. SP OnPrem / Subscription Edition (SE)

<p>On premises react app is deployed on SP onPrem server and utilizes SP List API [@microsoft/sp-http](https://www.npmjs.com/package/@microsoft/sp-http), the base communication layer for the SP REST services.  Authentication is not required since the SP site hosting the app handles access permissions.</p>

## Build / Package / Test

### spfx-webpart-list-app-ms365-onprem

#### sp-list-app

To start:
Go to sp-list-app directory
Run

```
npm install
```

Run

```
./removeMsalRefsThenGulpBuild.bash
```

**WARNING: Msal* files cause spfx-list-app bundling (gulp bundle) errors, but the Msal* services are not used**

Run

```
npm run build
```

Run

```
npm link
```

Run

```
npm start
```

#### spfx-list-app

to start developing against local server

Go to spfx-list-app-m365 directory
Run

```
npm install
```

Run

```
npm link sp-list-app
```

Run

```
gulp serve --nobrowser
```

to start developing against SP workbench

Deploy

```
gulp build
gulp bundle
gulp package-solution
```

Copy to app catalog
TODO more details...
