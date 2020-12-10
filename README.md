# MGT SPFx

Test project to address the issue with not being able to use multiple SPFx web parts on a page all using Microsoft Graph Toolkit.

Repo consists of an SPFx library component (mgt-spfx) meant to instantiate MGT components on the page and two tests project that use MGT to verify if the setup works.

## Prerequisites

- M365 tenant
- SharePoint app catalog configured to deploy the package
- admin account to deploy the packages
- Node v10
- Yeoman
- Gulp

## To test

- clone the repo
- build packages
  - build `mgt-spfx`
    - change directory to `mgt-spfx`
    - Restore dependencies: `npm i`
    - Build code: `gulp bundle --ship`
    - Create SPFx package: `gulp package-solution --ship`
    - Register package locally: `npm link`
  - build `mgt-wp-test1`
    - change directory to `mgt-wp-test1`
    - In `mgt-wp-test1/package.json` remove the reference to `mgt-spfx`
    - Restore dependencies `npm i`
    - Link `mgt-spfx`: `npm link mgt-spfx`
    - In `mgt-wp-test1/package.json` add back the reference to `mgt-spfx`
    - Build code: `gulp bundle --ship`
    - Create SPFx package: `gulp package-solution --ship`
    - repeat for `mgt-wp-test2`
- deploy packages
  - add the three .sppkg packages to SharePoint app catalog and deploy them to all sites
  - approve API permission requests for Graph in the SharePoint admin center at `https://[your-tenant]-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement`
- test web parts
  - navigate to the hosted SharePoint Framework workbench at `https://[your-tenant].sharepoint.com/_layouts/15/online/workbench.aspx`
  - add to the page the two MGT web parts