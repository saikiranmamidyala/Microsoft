# Exec Comms SharePoint/Teams app

Use `npm start` to dev

## SharePoint setup

- Content types
  - `Exec Event`
  - `Exec List`
- List
  - `Executives`
- Document library
  - `Executive Documents`
- Calendar
  - `Executive Events`
- Pages
  - `ConfigTab.aspx`
  - `RemoveTab.aspx`
  - `FutureEvents.aspx`
  - `PastEvents.aspx`
  - `Event.aspx`
  - `Assets.aspx`
  - `Calendar.aspx`
  - `Search.aspx`
  - `Permissions.aspx`
- Executive Groups
  - Name should be the same value as the `Title` field from the Executives list `+ " Members"`
  - Group owner should be
    - dev: `Exec Comms vNext Owners`
    - test: `Exec Comms vNext Dev Owners`
  - Who can view the membership of the group? `Group Members`
  - Who can edit the membership of the group? `Group Members`
  - Allow requests to join/leave this group? `No`
  - Choose the permission level group members get on this site: `Contribute`
  - *Create a new group*
    - *dev: https://blulinksolutions.sharepoint.com/sites/ExecCommsvNextDev/_layouts/15/newgrp.aspx?Source=https%3A%2F%2Fblulinksolutions%2Esharepoint%2Ecom%2Fsites%2FExecCommsvNextDev%2F_layouts%2F15%2Fuser%2Easpx*  
    - *test: https://blulinksolutions.sharepoint.com/sites/ExecCommsvNext/_layouts/15/newgrp.aspx?Source=https%3A%2F%2Fblulinksolutions%2Esharepoint%2Ecom%2Fsites%2FExecCommsvNext%2F_layouts%2F15%2Fuser%2Easpx*

## SPFx config

Package config: */config/package-solution.json*
- dev
  - id should end with `dd`
  - name: `exec-comms-dev`
  - zippedPackage: `solution/exec-comms-dev.sppkg`
- test
  - id should end with `00`
  - name: `exec-comms-test`
  - zippedPackage: `solution/exec-comms-test.sppkg`
-production
  -id should end with `00`
  -name `execcomms`
  -zippedPackage: `solutions/execcomms.sppkg`

Manifest config: */config/write-manifests.json*
- dev
  - cdnBasePath: `""` (empty string will resolve to localhost automatically for local dev)
- test
  - cdnBasePath: `https://blulinksolutions.sharepoint.com/sites/ExecCommsvNext/App/`

WebPart configs: */src/webparts/xxx/xxxWebPart.manifest.json*
- dev
  - id should end with `dd`
- test
  - id should end with `00`

## Steps to deploy for dev (only do the steps below once)
1. Update `package-solution.json`, `write-manifests.json`, and `*.manifest.json` for each web part (and the `SPFx config` section above shows how to do it).
2. `gulp clean`
3. `gulp bundle`
4. `gulp package-solution`
5. Drag and drop `exec-comms-dev.sppkg` from the `/sharepoint/solution` folder to our App Catalog https://blulinksolutions.sharepoint.com/sites/appcatalog/AppCatalog
6. When prompted, `Overwrite` the existing app.
7. You'll be prompted again. Click `Deploy` to trust the new app.

## Steps to deploy for test (repeatedly)
1. Update `package-solution.json`, `write-manifests.json`, and `*.manifest.json` for each web part (and the `SPFx config` section above shows how to do it).
2. `gulp clean`
3. `gulp bundle --ship`
4. Drag and drop all files in the `/temp/deploy` folder into our app folder in SharePoint https://blulinksolutions.sharepoint.com/sites/ExecCommsvNext/App
5. `gulp package-solution --ship`
6. Drag and drop `exec-comms-test.sppkg` from the `/sharepoint/solution` folder to our App Catalog: https://blulinksolutions.sharepoint.com/sites/appcatalog/AppCatalog
7. When prompted, `Overwrite` the existing app.
8. You'll be prompted again. Click `Deploy` to trust the new app.

## Steps to refer the icons and images from CDN location
1. Upload the `images` folder located in `src` folder onto CDN document library 
2. Update the `cdnAssetsBaseUrl` in `SharePoint.ts` file located under `src\shared` folder with the CDN document library Public CDN URL.

e.g.
https://publiccdn.sharepointonline.com/shrudomain.sharepoint.com/teams/C2CAppCDN/ExecComms/ExecComms

NOTE: Do not end the `cdnAssetsBaseUrl` with `/`