# Power BI Report Viewer WebPart or Teams Tab

Sample SPFX WebPPart to show Power BI reports, which can be deployed to SharePoint Online or Microsoft Teams

## Commands to run
```
npm install
gulp bundle
gulp serve
```

## Commands to debug
```
gulp serve --nobrowser
```
Then F5 or run Workbench debug option from the debugger of VS Code

## Command to bundle app to make avaliable to M365
```
gulp bundle --ship
gulp package-solution --ship
```
Then upload the app package file to SharePoint App Catelog

## References
- https://www.youtube.com/watch?v=Rh9pDG8kdX4
- https://www.youtube.com/watch?v=OcYHf3s_qH0
- https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant
- https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview
- https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis
- https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview
- https://aka.ms/m365pnp
