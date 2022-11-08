<h1>Sharepoint</h1>

- [Sharepoint](#sharepoint)
  - [Notes](#notes)
  - [References](#references)
- [Sharepoint Framework](#sharepoint-framework)
  - [Notes](#notes-1)
  - [Steps](#steps)
  - [References](#references-1)
- [PowerApps](#powerapps)
  - [Notes](#notes-2)
    - [Canvas App](#canvas-app)
    - [Common Data Service](#common-data-service)
    - [Portals](#portals)
  - [References](#references-2)
- [Power Automate AKA Flow](#power-automate-aka-flow)
  - [Notes](#notes-3)
  - [References](#references-3)

# Sharepoint
## Notes

## References

# Sharepoint Framework
## Notes
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


## Steps
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
5. Run the below command to start the application without a browser
    ```
    gulp serve --nobrowser
6. In order to create additional web parts, just run the command in the same path where we ran this command before. It will know whether to create a new solution or to add a new web part.
    ```
    yo @microsoft/sharepoint
7. Run the below command to compile the solution
    ```
    gulp build
8. Run the below command to minify the JavaScript files
    ```
    gulp bundle
9. Run the below command to create sharepoint package
    ```
    gulp package-solution
10. To install sppnpjs library, run the following command
    ```
    npm install sp-pnp-js --save
   
11. To debug custom library, link it using the below command
    ```
    npm link
12. use the below command to use the above custom library
    ```
    npm link <<customLibraryName>>
13. To install JQuery, run the below command
    ```
    npm install --save jquery
14. To install types for jquery, run the below command
    ```
    npm install --save @types/jquery
15. To install types for jquery ui, run the below command
    ```
    npm install --save @types/jqueryui
16. To install types for sharepoint graph api, run the below command
    ```
    npm install --save-dev @microsoft/microsoft-graph-types
17. Install azure cli
18. Run the below npm command 
    ```
    npm install --save azure-functions-ts-essentials
19. Run the below npm command 
    ```
    npm i -g azure-functions-core-tools@2 --unsafe-perm true
20. Run the below npm command
    ```
    npm install -g typescript
21. Login to azure portal
    1. Click on Login create a resource
    2. search for function app
    3. Create a function app
    4. Create a local function app in the local system by running the command `func init <function app name in azure portal>`
    5. run the command `func new`
    6. Select the template &rarr; Http Trigger
    7. Enter the function name 
    8. Go to the function folder
    9. Copy the relevant files from transpiled and project folders
    10. Start the function locally by running the command `func host start`
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
    ```

## References
* https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview
* https://developer.microsoft.com/en-us/fluentui#/styles/web/icons
* https://www.base64-image.de/
* https://pnp.github.io/pnpjs/getting-started/
* https://jsonplaceholder.typicode.com/
* http://json2ts.com/
* https://developer.microsoft.com/en-us/graph/graph-explorer
* https://aad.portal.azure.com

# PowerApps 
## Notes
* bedaye8031@seinfaq.com, test@123
* In Excel, create the data as a table.
### Canvas App
* Gallery 
    * Simple list of records pulled from a table.
    * Mostly a preview view with some data.
* Forms and Data cards
* Triggers 
    * Triggers start with the word "On"
    * SubmitForm &rarr; Submits the form
    * ResetForm &rarr; Resets the form
    * Back &rarr; Goes back
    * The functions are seperated by ";"
* The microsoft powerapp remembers the last screen we had. that's how back works.
* App checker can be used to identify errors and resolve it.
* To publish an app, save it and then publish it.
* For web app, use tablet layout if you are creating a blank app.
* Create a master screen where other screens are based off from. So all the screens will be identical with same color schemes.
* Wrap Count can be used to make a gallery as multicolumn gallery.
* There are 3 types of variables
    * Contextual variables
        * Can be accessed only within a screen.
        * To update a variable: UpdateContext({FirstVaraible: text1.text})
        * Can store many different items
    * Global variables
        * Can be set in one screen and accessed in another.
        * To update a global variable: Set(FirstVariable, text1.text)
        * Can store only one item
    * Collections
        * Data stored like excel
        * Add to a collection using Collect(OurCollections, {First: "Aroan", Second: "Kingslin"})
        * Remove from collection using Remove(OurCollections,ThisItem)
* Search takes in the search term whereas filter takes in the formula.
    * Search(Sheet1,TextInput5.Text,"FirstName","LastName","AgentName")
    * Filter(If(Dropdown1.Selected.Value = "All", true, VIPLevel = Dropdown1.Selected.Value))
* NewForm(Form) and EditForm(Form) is used to load the form in new or edit mode.
* Launch is used to open other apps like phone, email.
    * tel: 
    * mailto:
    * /providers/Microsoft.PowerApps/apps/{appId}
* Dropdown comparison with table values
    * SortByColumns(Distinct(Filter(Sheet1,Location = Dropdown2.Selected.Result),VIPLevel),"Result")
* Delete a record
    * Remove(Sheet1,Gallery1.Selected)
* To get selected date in datepicker, use Datepicker1.SelectedDate
* To get value from toggle, use Toggle2.value
* cons
    * Power apps has to be installed.
    * Power app users should be present in that org to access them.

### Common Data Service
* Benefits
    * All the data are standardized.
    * Business rules can be applied on the data.
    * Security and data isolation for each Users.
    * Data backups.
* Entities is now tables
    * Fields are columns
    * Relationships
        * Many to one
        * One to many
        * Many to many
    * Business Rules
    * Forms
    * Dashboards
    * Charts
* Business Rules
* Data flows
* Connections
* Gateways

### Portals
* Portals extends power apps as websites.
* Allows to share data external to the organization.
* To create a portal, click blank app > Blank Website > 

## References
* https://powerusers.microsoft.com/t5/Community-App-Samples/bd-p/AppFeedbackGallery
* https://make.powerapps.com/
* https://docs.microsoft.com/en-us/powerapps/powerapps-overview
* https://docs.microsoft.com/en-us/power-apps/maker/canvas-apps/working-with-formulas
* https://forwardforever.com/power-apps-and-git-version-control/
* https://powerapps.microsoft.com/en-us/pricing/


# Power Automate AKA Flow
## Notes
* Flow is a visual representation of any task. Trigger &rarr; Action.
* Flows can have multiple actions. It can be conditional actions.
* Types of Flows:
    
    | Type            | Trigger           | Description                                            |
    | --------------- | ----------------- | ------------------------------------------------------ |
    | Instant         | Button Click      |                                                        |
    | Automated       | An Event          | Automates any task                                     |
    | Scheduled       | A reoccuring time |                                                        |
    | Desktop Flow    | Any of the above  | Records and automates a process in the desktop/website |
    | Process Advisor | Any of the above  | Guides a user through a multistep process              |

* To use flows in power apps, **create it from the power apps page** and **not from power automate**.
    * Use actions in power app.
    * Use "Ask in Power Apps" in the flow.
    * If the flow is changed, the flow under actions has to refreshed or removed and added again.
* Create Http request triggers for third party application triggers.
* Create Parallel branches for parallel execution of actions.
* Approvals are manual interventions to actions.
    * **Create an approval**: Create an approval and proceeds to the next action without waiting.
    * **Start and wait for an approval**
    * **Wait for an approval**: We havn't created any approval so nothing to wait for.
* Outcome of an approval is "Approve".
* You can order/ move only actions that has the required fields filled out.
* You can multiple options in approvals.
* Terminate forces the flow to error out.
* "Apply to each" needs 2 arguments. One is Array and another is action. Current Item refers to the item from the loop.



## References
* https://make.powerautomate.com/home?fromflowportal=true
* https://learn.microsoft.com/en-us/power-automate/
* https://learn.microsoft.com/en-us/azure/logic-apps/workflow-definition-language-functions-reference


