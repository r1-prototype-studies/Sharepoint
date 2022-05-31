<h1>Sharepoint</h1>

- [Notes](#notes)
- [References](#references)

# Notes
* Creating self signed certificate is a one time activity. We need not create for all the web parts
* Folder Anatomy
    
    | Folder Name | Sub Folder  | Purpose                                                                                                                                      |
    | ----------- | ----------- | -------------------------------------------------------------------------------------------------------------------------------------------- |
    | Config      |             | All the config files are present here                                                                                                        |
    |             | Config.json | Has the web parts information along with the entrypoint and manifest. Has External library like Jquery information. Has localized resources. |
    | src         |             | Code related files are present here                                                                                                          |
    | teams       |             | tabs in microsoft Teams                                                                                                                      |
* If we are using skip feature deployment (set as true), deploy the sharepoint package in the app catalog. It will ask whether it needs to be available for all the sites, if selected yes, it will be available for all the sites.
* If we are not using skip feature deployment (set as false), deploy the sharepoint package in the app catalog. It will neither ask whether it needs to be available for all the sites and it is not available for all the sites. In order to make it available for all sites, follow the below steps:
  *  Deploy the package to the app catalog 
  *  We should install an app in the site collections. Go to settings and select "Add an app" 
  *  Click on package (app) that we uploaded.
* Import the below modules to make REST API calls
    ``` javascript
    import { 
        SPHttpClient,
        SPHttpClientResponse
    } from '@microsoft/sp-http'
* If you make any changes to the manifest.json file, you should stop the gulp server and start for the changes.
* The best place to assign default values is within init function.
* To disable reactive changes from propertypane, add the below method, by default this value is false.
    ``` javascript
    protected get disableReactivePropertyChanges(): boolean {
      return true;
    }
* To change the icon of the web part, change "officeFabricIconFontName" in the manifest file for an icon from office (refer link#1 in references) or use iconImageUrl for an image from internet.
* If we are working with the xml file, ensure that there are no spaces at the top.
* To create customized header and footer components, use extension.
* Extension uses: 
  * Application cutomizer - header / footer
  * Field Customizer - Graph color depending upon the status number
  * Command set - Custom command buttons
* In react, inline styles are not allowed.
* React lifecycle
  * Constructor
  * componentWillMount
  * Render
  * componentDidMount
* Don't give alias for jqueryUI. It already has a namespace. (JQueryUI)
* Sharepoint context is present in @microsoft/sp-webpart-base
* If the web part requires permission to access any backend api or Graph, it is recommended to have that web part as isolated web part.
  * This is to handle security concerns
  * Runs on a unique domain and is hosted on an iframe 
  * Permission granted only applies to that code running on that unique domain 
  * A dedicated azure AD registration gets created for this SPFx solution which handles the authentication.
* Extension cannot be an isolated because extension runs on the entire web page whereas isolated runs within an iframe.
* To make a webpart as an isolated webpart, update isDomainIsolated as true in package-solution.json
* To make a webpart as a provider webpart, import the below
    ``` javascript
    import {
        IDynamicDataPropertyDefinition,
        IDynamicDataCallables
    } from '@microsoft/sp-dynamic-data';
   ```
   * The provider webpart should implement the interface IDynamicDataCallables as like in 
   ``` javascript
   export default class ProviderWebPartDemoWebPart 
    extends BaseClientSideWebPart <IProviderWebPartDemoWebPartProps> 
    implements IDynamicDataCallables
    {
    ```
* Whenever the provider webpart communicates with the consumer webpart, the consumer webpart should have the below property of the type `DynamicProperty`
* Import the below in consumer webpart. 
    ``` javascript 
    import {   
        DynamicDataSharedDepth,
        PropertyPaneDynamicFieldSet,
        PropertyPaneDynamicField,
        IPropertyPaneConditionalGroup,
        IWebPartPropertiesMetadata
        } from '@microsoft/sp-webpart-base';

    import { DynamicProperty } from '@microsoft/sp-component-base';
* Run the below command to login to AWS from azzure cli
    ```
    az login
* To transpile a project manually, run the below command where tsconfig.json is present.
    ```
    tsc -p ./
* Use AADHttpClientFactory if we are calling azure ad secured rest apis.
* To make a webpart as a SPA, add "SharePointFullPage" to the supportedhosts in the manifest json file.
* TO load the web part with different locale, just add "--locale=es-es" to the gulp serve command.
* Run the below command to get the list of gulp tasks
    ```
    gulp --tasks
* To sequence the gulp tasks
    ``` javascript
    gulp.task('all-in-one-go', gulpSequence('clean', 'build', 'bundle', 'package-solution'));
* Adding sub task to a gulp task
    ``` javascript
    const subtaskbuildChild2 = build.subTask('sub-task-buildChild2', function (gulp, buildOptions, done) {
        this.log('sub-task-buildChild2 of build through this.log');
        done();    
    });

    build.task('sub-task-buildChild2', subtaskbuildChild2);

    build.initialize(gulp);

    if (gulp.tasks['build']) {
        gulp.tasks['build'].dep.push('sub-task-buildChild1','sub-task-buildChild2');
    }
* Adding pre and post build tasks
    ``` javascript
    const postBundlesubTask = build.subTask('post-bundle', function (gulp, buildOptions, done) {
        this.log('Message from Post Bundle Task');
        done();
    });
    build.rig.addPostBundleTask(postBundlesubTask);

    const preBuildSubTask = build.subTask('pre-build', function (gulp, buildOptions, done) {
        this.log('Message from PreBuild Task');
        done();
    });
    build.rig.addPreBuildTask(preBuildSubTask);

    const postBuildSubTask = build.subTask('post-build', function (gulp, buildOptions, done) {
        this.log('Message from PostBuild Task');
        done();
    });
    build.rig.addPostBuildTask(postBuildSubTask);

    const postTypeScriptSubTask = build.subTask('post-typescript', function (gulp, buildOptions, done) {
        this.log('Message from PostTypeScript task');
        done();
    });
    build.rig.addPostTypescriptTask(postTypeScriptSubTask);


# Steps
1. Run the below command to create the sharepoint project in VS code
    ```
    yo @microsoft/sharepoint
    ```
    To prevent installing the dependencies, run the following command instead
    ```
    yo @microsoft/sharepoint --skip-install
    ```
    To work with sharepoint beta version, run the following command
    ```
    yo @microsoft/sharepoint --plusbeta
2. Run the below command to create the self signed certificate 
    ```
    gulp trust-dev-cert
3. Open the file ./config/serve.json and change "enter-your-sharepoint-site" to the tenant we are using
4. Run the below command to start the application
    ```
    gulp serve
4. Run the below command to start the application without a browser
    ```
    gulp serve --nobrowser
4. In order to create additional web parts, just run the command in the same path where we ran this command before. It will know whether to create a new solution or to add a new web part.
    ```
    yo @microsoft/sharepoint
5. Run the below command to compile the solution
    ```
    gulp build
6. Run the below command to minify the JavaScript files
    ```
    gulp bundle
7. Run the below command to create sharepoint package
    ```
    gulp package-solution
8. To install sppnpjs library, run the following command
   ```
   npm install sp-pnp-js --save
9. To debug custom library, link it using the below command
    ```
    npm link
10. use the below command to use the above custom library
    ```
    npm link <<customLibraryName>>
1. To install JQuery, run the below command
    ```
    npm install --save jquery
2. To install types for jquery, run the below command
    ```
    npm install --save @types/jquery
3. To install types for jquery ui, run the below command
    ```
    npm install --save @types/jqueryui
1. To install types for sharepoint graph api, run the below command
    ```
    npm install --save-dev @microsoft/microsoft-graph-types
1. Install azure cli
2. Run the below npm command 
    ```
    npm install --save azure-functions-ts-essentials
3. Run the below npm command 
   ```
   npm i -g azure-functions-core-tools@2 --unsafe-perm true
4. Run the below npm command
    ```
    npm install -g typescript
5. Login to azure portal
    1. Click on Login create a resource
    2. search for function app
    3. Create a function app
    4. Create a local function app in the local system by running the command `func init <function app name in azure portal>`
    5. run the command `func new`
    6. Select the template &rarr; Http Trigger
    7. Enter the function name 
    8. Go to the function folder
    9. Copy the relevant files from transpiled and project folders
    9. Start the function locally by running the command `func host start`
1. Install office client cli using the below command
    ```
    npm install -g @pnp/office365-cli
2. Run the below command to grant access in office365 cli
    ```
    spo serviceprincipal grant add --resource "Microsoft Graph" --scope "Calendars.Read"
1. To work with team webpart creation, run the below command
    ```
    npm install @microsoft/teams-js --save
1. Install Sharepoint online management shell - https://www.microsoft.com/en-us/download/details.aspx?id=35588
2. Run the below commands in the Sharepoint online management shell
    ```
    Connect-SPOService -Url https://contoso-admin.sharepoint.com
    Set-SPOTenantCdnEnabled -CdnType Public
    Get-SPOTenantCdnEnabled -CdnType Public
    Get-SPOTenantCdnOrigins -CdnType Public
    Get-SPOTenantCdnPolicies -CdnType Public
3. To deploy in site Assets
    ```
    New-SPOPublicCdnOrigin -Url https://kameswarasarma.sharepoint.com/SiteAssets/HelloWorldWPCDN
4. Get the ID which you get from the below command and append it to https://publiccdn.sharepointonline.com/kameswarasarma.sharepoint.com/
    ```
    Get-SPOPublicCdnOrigins | Format-List
5. Open write-manifests.json file within config folder and update the value for cdnBasePath with that above url appended with ID.
1. Run the below command to install gulp sequence
    ```
    npm install --save-dev gulp-sequence




# References
* https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview
* https://developer.microsoft.com/en-us/fluentui#/styles/web/icons
* https://www.base64-image.de/
* https://pnp.github.io/pnpjs/getting-started/
* https://jsonplaceholder.typicode.com/
* http://json2ts.com/
* https://developer.microsoft.com/en-us/graph/graph-explorer
* https://aad.portal.azure.com




