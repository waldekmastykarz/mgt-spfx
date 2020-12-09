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
    - Restore dependencies `npm i`
    - Build code: `gulp bundle --ship`
    - Create SPFx package: `gulp package-solution --ship`
    - repeat for `mgt-wp-test2`
- deploy packages
  - deploy the three .sppkg packages to SharePoint app catalog
  - approve API permission requests for Graph