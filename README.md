# Sharepoint React App for SP Online / MS 365 and SP OnPrem and Subscription Edition (SE)

Example SP react app that reads list and adds items to list. The app's two variants:

1.  SP OnLine SPFx react webpart

    Webpart uses @microsoft/sharepoint-generator but seperates the react app which provides ability to choose react variants different than dependency in @microsoft/sharepoint-generator. Requires [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) in order to access SP Online site's information, specifically list data in this example app.

2.  SP OnPrem / Subscription Edition (SE)

    On premises react app is deployed on SP onPrem server and utilizes SP List API [@microsoft/sp-http](https://www.npmjs.com/package/@microsoft/sp-http), the base communication layer for the SP REST services. Authentication is not required since the SP site hosting the app handles access permissions, however the SP site digest must be included in the POST header for create, update, and delete list operations. The context API is:

    **${SITE_URL}/\_api/contextinfo**

## Prerequisites

### Required tools

- node v18
- gulp-cli
- yo
- Microsoft SharePoint Generator v1.18.0

### Required environments

- SP Online Tenant, [Set up MS 365 Tenant](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- SP On Premise, [Azure Account](https://azure.microsoft.com/en-us/free/search/?ef_id=_k_Cj0KCQjw7JOpBhCfARIsAL3bobcCY3SsETB1kTZWqsEQd0D0SYu2-uLikiVmjAyAfAtXtbNI_Poot6QaAs3ZEALw_wcB_k_&OCID=AIDcmm5edswduu_SEM__k_Cj0KCQjw7JOpBhCfARIsAL3bobcCY3SsETB1kTZWqsEQd0D0SYu2-uLikiVmjAyAfAtXtbNI_Poot6QaAs3ZEALw_wcB_k_&gad=1&gclid=Cj0KCQjw7JOpBhCfARIsAL3bobcCY3SsETB1kTZWqsEQd0D0SYu2-uLikiVmjAyAfAtXtbNI_Poot6QaAs3ZEALw_wcB) with Miscorosft Windows Server running SharePoint 2019

## node version 18

Recommend installing nvm for windows hosts [Node Version Manager](https://github.com/coreybutler/nvm-windows/releases), utility to switch node versions.

Node version 18 per @microsoft/generator-sharepoint

```
{
    name: "@microsoft/generator-sharepoint",
    version: "1.18.0",
    descriptions: "Yeoman generator for the SharePoint Framework",
    engines: {
        node: ">=16.13.0 <17.0.0 || >=18.17.1 <19.0.0"
    }
}
```

As **windows admin**:

```
nvm install 18
```

## gulp-cli

As **windows user**:

```
npm install --global gulp-cli
```

## yo

As **windows user**:

```
npm install --global yo
```

## SharePoint Generator

SP Generator [@microsoft/generator-sharepoint](https://www.npmjs.com/package/@microsoft/generator-sharepoint?activeTab=versions)

![version](https://img.shields.io/badge/version-1.18.0-green.svg)

As **windows user**:

```
npm install ---global @microsoft/generator-sharepoint@1.18.0
```

# Create SPFx Webpart Application / Bootstrap SP Webpart

Create SPFx webpart using SP Generator [SharePoint Framework](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview).
As **windows user**:

```
mkdir sp-list-app-m365-onprem
cd sp-list-app-m365-onprem
mkdir spfx-list-app
cd spfx-list-app
yo @microsoft/sharepoint@1.18.0
```

Recommend using MS command prompt since bash shell inputs are not as clear.

Bootstrap script inputs:

![SPFx script inputs](/images/spfx.png)

Bootstrap script finished:

![SPFx script finished](/images/spfxDone.png)

The **SP OnLine SPFx react webpart, variant 1** contains the code so the webpart can be deployed on a SP Online server. The SpfxListAppWebPart.ts wrapper adds the MS Graph API to the webpart context's http client and intializes the standalone react app with the SP site's data provider from its context.

The standalone react app does have a gulpfile dependency, [sp-build-web](https://www.npmjs.com/package/@microsoft/sp-build-web), necessary for variant 1 - build target that will run in web browser hosted from SP Online.

# Standalone React App (SP OnPrem)

The standalone react app defines all of the application features. This single app can be built for **SP OnPrem / Subscription Edition (SE), variant 2** or SP Online, variant 1.

The standalone react app does not have the same react version restrictions as the Microsoft SharePoint Generator utility. The only SharePoint dependency is **@microsoft/sp-http**, SharePoint REST services API for SP Lists. The **sp-build-web** gulp file dependency is only necessary for packaging the webpart (not needed for deploying the standalone react app on SP OnPrem or Subscription Edition/SE).

Standalone app: **sp-list-app**

## Initial Build

Switch to standalone app dir and install dependencies:

```
cd sp-list-app
npm install
```

## Build and Package

Note: Must **_npm install_** if its **package.json** is updated after initial build.

Set SharePoint environment for build script:

```
cd sp-list-app
gulp set-sp-site --site <SP Site> --listid <List ID> --graph <MS Graph URL>
```

Notes:

- SP Site => online tenant example: \*.sharepoint.com, onPrem example: soceur.\*/ppws/sanbox/ReactApps
- List Id => SP site -> navigate to list -> settings -> list settigs -> RSS settings
- MS Graph URL => _*graph.microsoft.com*_ for commercial

Remove Msal\* lib dependencies and then build with gulp (**must run from bash shell**):

```
./removeMsalRefsThenGulpBuild.bash
```

**WARNING: Msal-files cause spfx-list-app bundling (gulp bundle) errors, but the Msal-X services are not used.**

Bundle standalone app:

```
npm run build
```

Link standalone app libs:

```
npm link
```

The **npm link** feature copies the standalone app package to node's global node*modules -- \*\*\_only required for SPFx Webpart App*\*\*.

# SPFx Webpart App (SP Online)

SPFx webpart app: **spfx-list-app**

## Initial Deploy

Must delete dir and contents of **spfx-list-app/src/webparts/loc** and replace referenced variables in **SpfxAppWebPart.ts** with string constants.

Update **package-solution.json**, add MS Graph API permissions request (**webApiPermissionRequest**) for all Sites on tenant:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
  "solution": {
    "name": "spfx-list-app-client-side-solution",
    "id": "a0c9801e-fa75-4618-8bc9-992ade22b3b4",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "skipFeatureDeployment": true,
    "isDomainIsolated": false,
    "developer": {
      "name": "",
      "websiteUrl": "",
      "privacyUrl": "",
      "termsOfUseUrl": "",
      "mpnId": "Undefined-1.18.0"
    },
    "metadata": {
      "shortDescription": {
        "default": "spfx-list-app description"
      },
      "longDescription": {
        "default": "spfx-list-app description"
      },
      "screenshotPaths": [],
      "videoUrl": "",
      "categories": []
    },
    "webApiPermissionRequests": [
      {
        "resource": "Microsoft Graph",
        "scope": "Sites.ReadWrite.All"
      }
    ],
    "features": [
      {
        "title": "spfx-list-app Feature",
        "description": "The feature that activates elements of the spfx-list-app solution.",
        "id": "682fd9a9-94c5-41b9-b8ad-74a9025c9c91",
        "version": "1.0.0.0"
      }
    ]
  },
  "paths": {
    "zippedPackage": "solution/spfx-list-app.sppkg"
  }
}
```

Must add support for full-width column for webpart SP Online deploys, add **"supportsFullBleed": true,** to **SpfxListAppWebPart.manifest.json**. Updated file:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json",
  "id": "d0ae4c5b-968c-4f49-a2f7-453d3f33cd11",
  "alias": "SpfxListAppWebPart",
  "componentType": "WebPart",

  // The "*" signifies that the version should be taken from the package.json
  "version": "*",
  "manifestVersion": 2,

  // If true, the component can only be installed on sites where Custom Script is allowed.
  // Components that allow authors to embed arbitrary script code should set this to true.
  // https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f
  "requiresCustomScript": false,
  "supportedHosts": [
    "SharePointWebPart",
    "TeamsPersonalApp",
    "TeamsTab",
    "SharePointFullPage"
  ],
  "supportsThemeVariants": true,

  // Section layout, Full-width section for WebPart
  "supportsFullBleed": true,

  "preconfiguredEntries": [
    {
      "groupId": "5c03119e-3074-46fd-976b-c60198311f70", // Advanced
      "group": { "default": "Advanced" },
      "title": { "default": "spfx-list-app" },
      "description": { "default": "spfx-list-app description" },
      "officeFabricIconFontName": "Page",
      "properties": {
        "description": "spfx-list-app"
      }
    }
  ]
}
```

Switch to spfx app dir and install dependencies:

```
cd spfx-list-app
npm install
```

## Build and Package

The **"version"** key in **package-solution.json** defines the **App version** for the **Apps for SharePoint / App Catalog** when deployed to SP Online site, e.g. from **"1.0.0.0"** to **"1.0.0.1"**.

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
  "solution": {
    "name": "spfx-list-app-client-side-solution",
    "id": "a0c9801e-fa75-4618-8bc9-992ade22b3b4",
    "version": "1.0.0.1",
    ...
```

Note: Install not needed unless package.json is updated after initial install.

Add standalone app lib to local node_modules

```
npm link sp-list-app
```

Build for deploy to SP Online tenant:

```
gulp build
gulp bundle --ship
gulp package-solution --ship
```

Copy SPFx webpart to SP Tenant's AppCatalog, URL:

**${SITE_NAME}.sharepoint.com/sites/appcatalog/AppCatalog/Forms/AllItems.aspx**

Steps:

- Click on - Distribute apps for SharePoint
- Drag and drop **spfx-list-app/sharepoint/solution/spfx-list-app/sppkg** onto webpage
- Replace It
- Deploy - do **NOT** select **Make this solution available to all sites in the organization**
  App version will match package-solution.json "version".

As **SP Admin**, goto URL:
**${SITE_NAME}-admin.sharepoint.com/\_layouts/15/online/AdminHome.aspx#/home**

Note: Requires SP Admin assistance for our enterprise SP Online.

Click on:

- Advanced
- API Access

Grant access to **Microsoft Graph** in **Pending requests**.

API access updated:

![SP Online MS Graph API Access](/images/spOnlineAdminApiAccess.png)

# Deploy SPFx App

## Prerequisite

Must create **ListAppExample** SP list as defined in **IListItem** interface.

## Initial Deploy Steps

- Add new app, **add an app** from communication site's **Site Contents**
- Add **spfx-list-app-client-side-solution** under **Apps you can add**
- Add new page, **+ New -> Site Page** communication site's **Pages**
- Add **Full-width section**
- Add **spfx-list-app**
- Publish

## Update Deploy Steps

- Copy SPFx webpart to SP Tenant's AppCatalog
- Click webpart's **... -> ABOUT** under communication site's **Site contents**
- Click **GET IT**

Note: SP Page will update webpart automatically.

# Test SPFx Webpart Locally

Develop and test against SP workbench on tenant.

## Prerequisites

- Must deploy SPFx webpart to SP Online tenant and must grant API access to MS Graph before running in dev mode from localhost.
- Must create **ListAppExample** SP list as defined in **IListItem** interface.

## Build then Link sp-list-app

Build and copy standalone lib.

```
cd sp-list-app
gulp set-sp-site --site <SP Site> --listid <List ID>
./removeMsalRefsThenGulpBuild.bash && npm run build && npm link
```

Notes:

- SP Site => online tenant example: \*.sharepoint.com
- List Id => SP site -> navigate to list -> settings -> list settings -> RSS settings

## Link then build spfx-list-app

Run in development mode on tenant's workbench.

```
cd spfx-list-app
npm link sp-list-app && gulp serve --nobrowser
```

Open browser and navigate to URL:

**${SITE_NAME}.sharepoint.com/sites/dev/\_layouts/15/workbench.aspx**
