Module 04 - Hook into Office 365 and SharePoint with Single Page Apps
=====================================================================
In this lab, configure Apps for Office in Word and Outlook. The lab uses an existing application that can be found in the src folder.

##Overview
This lab enables students to take an existing ASP.NET application, configure it with an Azure AD app & create a single page application that reads and writes data to SharePoint in Office 365.

This lab uses an existing ASP.NET web application that can be found in the `/src/StarterProject` folder. To save some time with the coding portion of this lab, empty files have already been added that you will fill in with the code outlined in the steps below.

> If you are interested in seeing an advanced version of the sample project you build in this lab, refer to the [Research Project Code Sample](https://github.com/OfficeDev/Research-Project-Code-Sample) project on the [OfficeDev GitHub](https://github.com/OfficeDev) account. It adds some extra user interface controls like a busy animation & toast notifications as well as a master-detail implementation to the lab. 
> 
> These additional pieces would make this lab unnecessarily long and therefore have been omitted form the sample you will build in this lab.

##Objectives
- Learn to configure Microsoft Azure AD for an Office 365 app.
- Understand how and why you need to secure an OAuth access token.
- Create a single page app.

##Prerequisites
- Visual Studio 2013 for Windows 8
- You must have an Office 365 tenant and Windows Azure subscription to complete this lab.
- You must have completed the lab associated with Module 2.

##Exercises
The hands-on lab includes the following exercises:
- [Create and configure an Azure AD Application](#exercise1)
- [Explore and Configure the Starter Project](#exercise2)
- [Build the Single Page App](#exercise3)

<a name="exercise1"></a>
##Exercise 1: Create and configure an Azure AD Application
In this exercise you will create and configure an Azure AD application that the single page application will leverage.

###Task 1 - Create Azure AD Application
Follow these steps to create the Azure AD application.

1. Log into the [Azure Management Portal](https://manage.windowsazure.com) as an administrator of your subscription.
1. Click **Active Directory**.
1. Click the Azure Active Directory associated with your subscription.
1. Click **Applications**.
    ![Figure 1](img/01.png?raw=true)
1. Click **Add**.
1. Click **Add an application my organization is developing**.
    ![Figure 2](img/02.png?raw=true)
1. On the **Tell us about your application** screen, do the following:
    1. Set the **Name** to **SpResearchTrackerSPA**. 
    1. Set the **Type** to **Web Application and/or Web API**
    1. Click the **Arrow** at the bottom right of the screen.
    ![Figure 3](img/03.png?raw=true)
1. On the **App Properties** screen, do the following:
    1. Enter **https://localhost:44362/** in the **Sign-On URL** field.
    1. Enter **https://localhost:44362/SpResearchTracker** in the **App ID URI** field.
1. Click the **Checkmark** at the bottom right of the screen.
   ![Figure 4](img/04.png?raw=true)

Now you have created a new Azure AD app.

###Task 2 - Configure the Azure AD Application
Follow these steps to configure the Azure AD app. It is assumed you are currently on the page for the newly created app, which is where you are taken after creating the app. If you aren't, navigate to it in the Azure Management Portal.

1. Click **Configure**.
     ![Figure 5](img/05.png?raw=true)
1. Locate and copy the **Client ID**. **Save** the value for later.
     ![Figure 6](img/06.png?raw=true)
1. In the **Keys** section, select **2 Years** for the duration.
     ![Figure 7](img/07.png?raw=true)
1. Click the **Save** button in the footer of the page.
     ![Figure 8](img/08.png?raw=true)
1. After the settings save, copy the key. **Save** the value for later.
     ![Figure 9](img/09.png?raw=true)
1. In the **Reply URL** section, add **https://localhost:44362**
     ![Figure 10](img/10.png?raw=true)
1. Click the **Save** button in the footer of the page.
     ![Figure 11](img/08.png?raw=true)
1. In the **Permissions to Other Applications** section, do the following:
    1. Select **Office 365 SharePoint Online**.
    1. Grant **Create or Delete Items**, **Edit Items**, and **Read Items** permissions.
     ![Figure 12](img/11.png?raw=true)
1. Click the **Save** button.
     ![Figure 13](img/08.png?raw=true)

Now you have configured the Azure AD app.

<a name="exercise2"></a>
##Exercise 2: Explore and Configure the Starter Project
In this exercise you will explore the existing starter project provided. A starter project is provided in the interest of time.

You may elect to skip the first task in this exercise. It is provided to familiarize yourself with the starter project that you will use for the reminder of this lab.

###Task 1 - Explore the Starter Project 
In this task you will explore a starter project provided to save you some time. 

> *This starter project is an ASP.NET MVC application that includes all the code necessary to provide a RESTful OData service to for the single page application (SPA).*
> 
> *The SPA you will create will not communicate with Office 365 directly, rather it will go through this ASP.NET web application. This is done so the Office 365 OAuth2 access token never touches the client; it is retrieved by the ASP.NET application form Azure AD and stores it in a server-side session.*
>
> *When the SPA calls the ASP.NET application, it retrieves the access token from the session state and issues an HTTPS request to Office 365's REST API, including the access token in the request.*

1. Locate & open the starter solution **[SpAppResearchTracker.sln](src/StarterProject/SpAppResearchTracker.sln)** in Visual Studio 2013.
1. Open the file **App_Start\BundleConfig.cs**. Notice it has a few entries to load the included JavaScript and CSS files that are already part of the project.
1. Open the file **App_Start\RouteConfig.cs**. This is the MVC routing configuration file. It has only a single route defined to catch all inbound requests. Notice if nothing is specified, it defaults to the **Home** controller and **Index** view.
1. Open the file **Views\Shared\\_Layout.cshtml**. This is the shared MVC view use by the application. There is nothing special about this view at this time... just notice the line `@RenderBody()` as that is where the main view will be inserted.
1. The **Views\Home\Index.cshtml** view does not contain anything important, so you can skip that. However open the controller for this view, **Controllers\HomeController.cs** & jump to the `Index()` method.
    - This is the first page a user is taken to when navigating to the site.
    - It pulls some values from the `web.config` & creates a redirect URL, `redirectUri`, where the user should be taken to once they have successfully authenticated.
    - The `OAuthController` is then called which is a class provided by the [Microsoft Azure AD Samples & Documentation](https://github.com/AzureADSamples) team in project [WebApp-WebAPI-OAuth2-UserIdentity-DotNet](https://github.com/AzureADSamples/WebApp-WebAPI-OAuth2-UserIdentity-DotNet). 
      
        The `OAuthController.GetAuthorizationUrl()` method initiates the authentication process. This takes the user to Azure AD to login. Upon a successful authentication, the user is sent back to the site, specifically the path **/OAuth**.        

        This postback form the Azure AD authentication process is handled by the `OAuthController.Index()` method that will store the provided OAuth2 access token & refresh token in cache (*refer to lines [110](/src/StarterProject/SpResearchTracker/Controllers/OAuthController.cs#L110) & [111](/src/StarterProject/SpResearchTracker/Controllers/OAuthController.cs#L117)*) for later use.

        Finally the user is redirected back to the specified homepage of the site (*see [line 117](/src/StarterProject/SpResearchTracker/Controllers/OAuthController.cs#L117) in **OAuthController.cs***) that was specified at the start of the authentication process (*see line [19](src/StarterProject/SpResearchTracker/Controllers/HomeController.cs#L19) in the **HomeController.cs***).
1. After the user authenticates, they are taken to **/Home/SPA** which is rendered in the view **Views\Home\SPA.cshtml**.
    - Notice the code that is obtaining an anti-forgery token from the server when the view is rendered and written to a hidden form variable.
    - This will be included in the header of every AJAX request by the SPA and is used by the custom OData service to:
        1. Validate the user making the request has been authenticated and 
        1. Obtain the OAuth access token from the session state that will be included in the request to the Office 365 REST API.
1. The **Controllers\ProjecsController.cs** handles all OData HTTP calls by the SPA.
1. It uses the underlying **Models\ProjectsRepository.cs** class to create the HTTP requests that are sent to the Office 365 REST API & include the OAuth2 access token in the header setting **Authorization**.

In this task you opened and explored the starter project. You should also have an understanding how the authentication is setup and how the Office 365 OAuth2 access token obtained by the ASP.NET application is protected by never exposing it to the client and exclusively keeping it server-side.

###Task 2 - Configure the Starter Project for the Azure AD Application
In this task you will configure the starter project to use the Azure AD application created in the [first exercise](#exercise1).

1. Locate & open the starter solution **[SpAppResearchTracker.sln](src/StarterProject/SpAppResearchTracker.sln)** in Visual Studio 2013.
1. Open the **web.config** file.
    - **Replace** the **[[REPLACE]]** in the **ida:FederationMetadataLocation** setting with your Office 365 tenant ID.
    - **Replace** the **ida:Realm** setting with **APP ID URL** value you entered when creating the Azure AD app: **https://localhost:44362/SpResearchTracker**.
    - **Replace** the **ida:AudienceUri** setting with same value you used in the **ida:Realm** setting.
    - **Replace** the **ida:Audience** setting with same value you used in the **ida:Realm** setting.
    - **Replace** the **ida:Tenant** setting with the appropriate value for your Office 365 environment.
    - **Replace** the **ida:ClientID** setting with the value you saved earlier when creating the Azure AD application.
    - **Replace** the **ida:AppKey** setting with the value you saved earlier when creating the Azure AD application.
    - **Replace** the **ida:Password** setting with the same value you used in the **ida:AppKey** setting.
    - **Replace** the **ida:Resource** setting with the appropriate value for your environment (your Office 365 tenant ID).
    - **Replace** the **ida:SiteUrl** setting with the value of a SharePoint site collection in your Office 365 subscription.
        - Make sure you have permissions to create lists within the site you specify.
        - Make sure you **include** a trailing **/** in the URL.
    - Locate the `<system.identityModel>` node. Within this node you will find another node, `<audienceUris>`.
        - **Replace** the `value` attribute on the `<add>` element to the same URL for the SharePoint site you entered for the **ida:SiteUrl** setting above.
    - Locate the `<system.identityModel.services>` node. Within this node you will find another node, `<wsFederation>`
        - **Replace** the **[[REPLACE]]** in the `issuer` attribute with your Office 365 tenant ID.
        - **Replace** the **[[REPLACE]]** in the `realm` attribute with same value you used in the **ida:Realm** setting above.

> It may seem strange that you have to enter the same value a few times for different settings. This is because there are different 3rd party libraries that are used in the project that look for the same information using different settings values.

In this task you configured the starter project to use the Azure AD application created in the [first exercise](#exercise1). 


<a name="exercise3"></a>
##Exercise 3: Build the Single Page App
In this exercise you will finish the starter project to implement the single page application.

###Task 1 - Add Necessary NuGet Packages
In this task you will add a series of NuGet packages to the project that will provide the necessary libraries to create a single page application with [AngularJS](https://angularjs.org) as well as other dependencies.

> You will use the NuGet Package Manager Console to manually add packages with specific version numbers. Usually you can omit the version number and get the current version, but to ensure no new package versions break this lab, you will use the specific versions used in creating this lab.
>
> This task assumes you still have the **[SpAppResearchTracker.sln](src/StarterProject/SpAppResearchTracker.sln)** solution open in Visual Studio 2013 from the last task. If you don't, open it now.

1. In Visual Studio, use the menu to open the NuGet Package Manager Console: **Tools > NuGet Package Manager > Package Manager Console**.

1. The Package Manager Console will have a yellow status alert indicating some packages are missing. Click the **Restore** button on the right-hand side of the status message to download all the missing packages that are currently referenced in the solution.

1. Enter the following into the Package Manager Console to add **AngularJS.Core**:

    ````powershell
    PM> Install-Package -Id "AngularJS.Core" -Version 1.2.16
    ````

1. Enter the following into the Package Manager Console to add **AngularJS.Route**:

    ````powershell
    PM> Install-Package -Id "AngularJS.Route" -Version 1.2.16
    ````

1. Enter the following into the Package Manager Console to add **AngularJS.Sanitize**:

    ````powershell
    PM> Install-Package -Id "AngularJS.Sanitize" -Version 1.2.16
    ````

1. Enter the following into the Package Manager Console to add **AngularJS.Animate**:

    ````powershell
    PM> Install-Package -Id "AngularJS.Animate" -Version 1.2.16
    ````

1. Enter the following into the Package Manager Console to add **bootstrap**:

    ````powershell
    PM> Install-Package -Id "bootstrap" -Version 3.2.0
    ````

1. Enter the following into the Package Manager Console to add **Angular.UI.Bootstrap**:

    ````powershell
    PM> Install-Package -Id "Angular.UI.Bootstrap" -Version 0.11.0
    ````

1. Enter the following into the Package Manager Console to add **Breeze.Angular**:

    ````powershell
    PM> Install-Package -Id "Breeze.Angular" -Version 0.8.7
    ````

1. Enter the following into the Package Manager Console to add **Breeze.EdmBuilder**:

    ````powershell
    PM> Install-Package -Id "Breeze.EdmBuilder" -Version 1.0.5
    ````

In this task you added all the necessary NuGet packages to the project that you will use in the building of the SPA.

###Task 2 - Add References to JavaScript & CSS Libraries
In this task you will update the starter project to include references to all the JavaScript and CSS files added in the NuGet packages.

> This task assumes you still have the **[SpAppResearchTracker.sln](src/StarterProject/SpAppResearchTracker.sln)** solution open in Visual Studio 2013 from the last task. If you don't, open it now.

1. Open the **App_Start\BundleConfig.cs** file.
1. Locate the existing `StyleBundle` named **~/Content/css" in the **App_Start\BundleConfig.cs** file & replace it with the following code:

    ````csharp
    bundles.Add(new StyleBundle("~/Content/css").Include(
            "~/Content/bootstrap.css",
            "~/Content/bootstrap-theme.css",
            "~/Content/site.css",
            "~/Content/animation.css"));
    ````

1. Now, add the following code to the **App_Start\BundleConfig.cs** file to create a few more bundles. Refer to the comments in the code for an explanation on each one:

    ````csharp
    // bootstrap related libraries (form NuGet packages)
    bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
            "~/Scripts/bootstrap.js",
            "~/Scripts/respond.js"));

    // angular bundle (from NuGet packages)
    bundles.Add(new ScriptBundle("~/bundles/angular").Include(
            "~/Scripts/angular.js",
            "~/Scripts/angular-route.js",
            "~/Scripts/angular-sanitize.js",
            "~/Scripts/angular-animate.js",
            "~/Scripts/angular-ui/ui-bootstrap-tpls.js"));

    // breeze bundle (form NuGet packages)
    bundles.Add(new ScriptBundle("~/bundles/breeze").Include(
            "~/Scripts/datajs-1.1.3.js",
            "~/Scripts/q.js",
            "~/Scripts/breeze.debug.js",
            "~/Scripts/breeze.angular.js"));

    // spa bootstrapping bundle (empty JS files to be coded)
    bundles.Add(new ScriptBundle("~/bundles/appcore").Include(
            "~/App/app.js",
            "~/App/config.js",
            "~/App/config.route.js",
            "~/App/config.angular.js",
            "~/App/config.breeze.js"));

    // spa common modules (empty JS files to be coded)
    bundles.Add(new ScriptBundle("~/bundles/appcommonmodules").Include(
            "~/App/common/common.js"));

    // spa controllers (empty JS files to be coded)
    bundles.Add(new ScriptBundle("~/bundles/appcontrollers").Include(
            "~/App/dashboard/dashboard.js",
            "~/App/layout/shell.js",
            "~/App/projects/list.js",
            "~/App/projects/detailView.js", 
            "~/App/projects/detailEdit.js"));

    // spa services (empty JS files to be coded)
    bundles.Add(new ScriptBundle("~/bundles/appservices").Include(
            "~/App/services/datacontext.angular.js",
            "~/App/services/datacontext.breeze.js"));
    ````

    Save your changes and close the **App_Start\BundleConfig.cs** file.
1. Open the **Views\Shared\\_Layout.cshtml** file.
1. Before the closing `</head>` tag, add the following code to the **Views\Shared\\_Layout.cshtml** file to include all the new `ScriptBundles` defined in the last step:

    ````csharp
    @* angular modules *@
    @Scripts.Render("~/bundles/angular")
    @* breeze libraries *@
    @Scripts.Render("~/bundles/breeze")
    @* custom app modules, services and utilities *@
    @Scripts.Render("~/bundles/appcore")
    @Scripts.Render("~/bundles/appcommonmodules")
    @Scripts.Render("~/bundles/appcontrollers")
    @Scripts.Render("~/bundles/appservices")
    ````

1. Locate the closing `</html>` tag and add the following to include the bootstrap bundle to the the **Views\Shared\\_Layout.cshtml** file;

    ````csharp
    @Scripts.Render("~/bundles/bootstrap")
    ````

    Save your changes and close the **Views\Shared\\_Layout.cshtml** file.

In this task you updated the starter project to include references to all the JavaScript and CSS files added in the NuGet packages.

###Task 3 - Code the Single Page Application
In this task you will add code to existing JavaScript and HTML files to implement the single page application (SPA).

> This task assumes you still have the **[SpAppResearchTracker.sln](src/StarterProject/SpAppResearchTracker.sln)** solution open in Visual Studio 2013 from the last task. If you don't, open it now.

1. First update the shared MVC layout to initialize and bind the AngularJS app to the page.
    1. Open the **Views\Shared\\_Layout.cshtml** file.
    1. Add an Angular **app** directive to the opening `<html>` element:

        ````xml
        <html data-ng-app="app">
        ````

    1. Next, make the page title dynamic by adding an Angular **bind** directive to the page's `<title>` element:

        ````xml
        <title data-ng-bind="'Project Research Tracker' + pageTitle"></title>
        ````

    Save your changes and close the **Views\Shared\\_Layout.cshtml** file.
1. Create the Angular module that will serve as the single page application.
    1. Open the **App\app.js** file.
    2. Add the following code to the body of the JavaScript self-executing function: 
 
        ````javascript
        // create the app
        var app = angular.module('app', [
            // ootb angular modules
            'ngRoute',      // app routing support
            'ngSanitize',   // fixes html issues with some data binding
            'ngAnimate',    // adds animation capabilities

            'ui.bootstrap',

            // breeze modules
            'breeze.angular',     // wires up breeze with angular automatically

            // app modules
            'common'
        ]);

        // startup code - this runs before the app actually "starts"
        app.run(['$route', 'breeze.config', appStartup]);
        function appStartup($route, breezeConfig) {
            // $route - routes are loaded first
            // $breezeConfig - breeze global configuration module... triggers
            //    the initial load of the metadata from the WebAPI OData feed
            //    used throughout the app
        }
        ````

    Save your changes and close the **App\app.js** file.
1. You referenced two JavaScript modules in the above code that are not yet created: **common** & **breeze.config**. Let's address those now:
    1. Open the **App\common\common.js** file.
    1. Add the following code to the body of the JavaScript self-executing function:

        ````javascript
        // create module
        var commonModule = angular.module('common', []);
        // create provider
        commonModule.provider('commonConfig', function () {
          this.config = {};
          this.$get = function () {
            return {
              config: this.config
            };
          };
        });
        // create the common service
        commonModule.factory('common',
          ['$window', '$q', '$rootScope', '$timeout', 'commonConfig', common]);
        // create the factory 'common'
        function common($window, $q, $rootScope, $timeout, commonConfig) {
          // public signature of module
          var service = {
            // pass though common angular dependencies
            $broadcast: $broadcast,
            $q: $q,
            $timeout: $timeout,
            // my services
            activateController: activateController,
            // global util functions
            goBack: goBack
          };
          return service;
          // pass through of the angular $broadcast service
          function $broadcast() {
            return $rootScope.$broadcast.apply($rootScope, arguments);
          }
          // global function used to activate a controller once all promises have completed
          function activateController(promises, controllerId) {
            return $q.all(promises).then(function () {
              var data = { controllerId: controllerId };
              $broadcast(commonConfig.config.controllerActivateSuccessEvent, data);
            });
          }
          // navigate backwards in the history stack
          function goBack() {
            $window.history.back();
          }
        }
        ````

        Save your changes and close the **App\common\common.js** file.

    1. Open the **App\config.breeze.js** file. The SPA will leverage the JavaScript library [BreezeJS](http://www.breezejs.com) to simplify REST queries.
    1. Add the following code to the body of the JavaScript self-executing function:

        ````javascript
        // define service
        var serviceId = 'breeze.config';
        angular.module('app').factory(serviceId,
          ['$q', 'breeze', configBreeze]);
        // create service
        function configBreeze($q, breeze) {
          // breeze's entity manager, the 'center of gravity' for all 
          // breeze communication this will be shared across all modules
          var entityManager;
          // init service
          init();
          // service public signature
          return {
            entityManager: entityManager,
            dataService: getDataService()
          };
          function init() {
            // configure breeze to use WebAPIOData data service adapter & set as the default
            breeze.config.initializeAdapterInstance('dataService', 'webApiOData', true);
            // init entity manager
            initEntityManager()
              .then(function () {
                // ensure antiforgery token included in all HTTP requests
                configureDetaultHttpClient();
                // fixup the metadata processing of the response
                fixupMetadataProcessing();
              });
          }
          // create breeze data service w/ endpoint to the WebAPIOdata service
          function getDataService() {
            // set the data service endpoint
            return new breeze.DataService({
              serviceName: window.location.protocol + '//' + window.location.host + '/odata/'
            });
          }
          // initialize breeze entity manager
          function initEntityManager() {
            var deferred = $q.defer();
            // if entity manager hasn't been initialized, do it now
            if (entityManager === undefined) {
              // metadataStore must be populated before adding validators, so manually fetch metadata
              entityManager = new breeze.EntityManager(getDataService());
              // ^^^ normally this happens on the first query, but if someone jumps straight to
              //  a deep link to create something (ie: #/projects/new), no query will happen
              //  and therefore none of the validators will be created
              entityManager.fetchMetadata()
                .then(function () {
                  configEntityIdentity();
                  attachValidatorsToBreezeEntities();
                  deferred.resolve();
                })
                .catch(function (exception) {
                  deferred.reject();
                });
            }
            return deferred.promise;
          }
          // configure the default HTTP client to include the antiforgery token in all requests
          function configureDetaultHttpClient() {
            // breeze uses datajs for OData requests
            // get a copy of the default OData client datajs uses
            var unsecuredClient = OData.defaultHttpClient;
            // create a new 'secured' http client...
            var securedClient = {
              request: function (request, success, error) {
                // that adds the verification token to it
                request.headers.RequestVerificationToken = jQuery("#__RequestVerificationToken").val();
                return unsecuredClient.request(request, success, error);
              }
            };
            // configure datajs to use the new secured client
            OData.defaultHttpClient = securedClient;
          }
          // configure the settings for the ID field as an identity
          function configEntityIdentity() {
            // set the project entity ID settings
            var projectType = entityManager.metadataStore.getEntityType('Project');
            projectType.autoGeneratedKeyType = breeze.AutoGeneratedKeyType.Identity;
          }
          // add custom client-side validators to the breeze entities
          function attachValidatorsToBreezeEntities() {
            // project entity
            //  title = required field 
            entityManager.metadataStore.getEntityType('Project')
              .getProperty('Title').validators.push(breeze.Validator.required());
          }
          // fixes the metadata response to ensure that the etag is treated as real metadata
          function fixupMetadataProcessing() {
            // get reference to the base 'visitNode' method
            var adapter = breeze.config.getAdapter('dataService');
            var visitNodeBase = adapter.prototype.jsonResultsAdapter.visitNode;
            // create a new 'visitNode' method that looks for the etag as a property on the entity
            //  if it isn't present in the metadata      
            adapter.prototype.jsonResultsAdapter.visitNode = function (node, mappingContext, nodeContext) {
              var metadata = node && node.__metadata;
              if (metadata && !metadata.etag && node.__eTag) {
                metadata.etag = node.__eTag;
              }
              //  .. then call the base 'visitNode'
              return visitNodeBase(node, mappingContext, nodeContext);
            }
          }
        }
        ````

        Save your changes and close the **App\config.breeze.js** file.
    1. The Breeze library, when used with Angular, relies on the Angular $http service for it's AJAX calls to the server. To configure all calls to use the SharePoint JSON format response, configure the Angular $http service. 
        
        Open the **App\config.angular.js** file.
    1. Add the following code to the body of the JavaScript self-executing function:

        ````javascript
        // define service
        var serviceId = 'angular.config';
        angular.module('app').factory(serviceId,
          ['$http', configAngular])         
        // create service
        function configAngular($http) {
          // init factory
          init()            
          // service public signature
          return {}         
          // init factory
          function init() {
            // set common $http request headers
            $http.defaults.headers.common.Accept = 'application/json;odata=verbose;';
          }
        }
        ````

        Save your changes and close the **App\config.angular.js** file.
1. With the core SPA created, now configure it:
    1. Open the **App\config.js** file.
    1. Add the following code to the body of the JavaScript self-executing function:

        ````javascript
        // get a reference to the app module
        var app = angular.module('app');

        // collection of events used throughout the app
        var events = {
            // event when the controller has been successfully activated
            controllerActivateSuccess: 'controller.activateSuccess',
        };

        // create global static config object
        var config = {
            title: 'Project Research Tracker App',
            events: events
        };
        app.value('config', config);

        // setup the angular logging provider to debug=on
        app.config(['$logProvider', function($logProvider) {
            if ($logProvider.debugEnabled) {
              $logProvider.debugEnabled(true);
            }
        }]);

        // configure the common configuration
        app.config(['commonConfigProvider', function (cfg) {
            // setup app events
            cfg.config.controllerActivateSuccessEvent = config.events.controllerActivateSuccess;
        }]);
        ````

        Save your changes and close the **App\config.js** file.
    1. Now configure the SPA for all URL routes that are supported. Open the **App\config.route.js** file.
    1. Add the following code to the body of the JavaScript self-executing function:

        ````javascript

        // get a reference to the app module
        var app = angular.module('app');
        // put all routes into global constant
        app.constant('routes', getRoutes());
        // config the routes & their resolvers
        app.config(['$routeProvider', 'routes', routeConfigurator]);
        function routeConfigurator($routeProvider, routes) {
          // load all routes into the system
          routes.forEach(function (route) {
            $routeProvider.when(route.url, route.config);
          });
          // if a route isn't found, send to spa's 404
          $routeProvider.otherwise({ redirectTo: '/' });
        }
        // build routes
        function getRoutes() {
          return [
            {
              url: '/',
              config: {
                templateUrl: '/App/dashboard/dashboard.html',
                title: 'Dashboard'
              }
            },
            // projects
            {
              url: '/projects',
              config: {
                templateUrl: '/App/projects/list.html',
                title: 'Projects'
              }
            },
            {
              url: '/projects/:projectId',
              config: {
                templateUrl: '/App/projects/detailView.html',
                title: 'View an Existing Project',
                templateMode: 'view'
              }
            }
          ];
        }
        ````

        Save your changes and close the **App\config.route.js** file.

1. Now we need a core shell for our SPA where Angular will load the views.
    1. Open the **Views\Home\SPA.cshtml** file.
    1. Locate the hidden input form control on the page and add the following `<div>` just above it. It contains an Angular **include** directive which will replace the `<div>` at runtime:
    
        ````xml
        <div data-ng-include="'/App/layout/shell.html'"></div>
        ````
        
        Save your changes and close the **Views\Home\SPA.cshtml** file.

    1. Now create the shell view. Open the **App\layout\shell.html** file. It already has all the HTML it needs, but it needs some Angular directives to be functional.
    1. Locate the opening `<div>` and add the following Angular **controller** directive:
    
        ````xml
        <div data-ng-controller="shell as vm">
        ````
    
    1. Next, locate the `<div id="splash-page">` and add the **show** Angular directive so we can conditionally show/hide the splash page:
    
        ````xml
        <div id="splash-page" data-ng-show="vm.showSplashPage" class="dissolve-animation">
        ````

    1. Now implement the working animation message for the app. Locate the `<div class="page-splash dissolve-animation">` and add the following **show** directive:
    
        ````xml
        <div data-ng-show="vm.isBusy" class="page-splash dissolve-animation">
        ````

       For the inner empty `<div>`, add the following to bind the **busyMessage** property from the controller to the view:
    
        ````xml
        <div class="page-splash-message page-splash-message-subtle">{{vm.busyMessage}}</div>
        ````

    1. Finally, update the main view container. Find the empty `<div>` after the **main view container** commend replace it with the following. When Angular finds the matching route, it will load the view in the **view** directive:
    
        ````xml
        <div data-ng-view class="shuffle-animation"></div>
        ````
        
        Save your changes and close the **App\layout\shell.html** file.

    1. With the shell view complete, implement the controller. Open the **App\layout\shell.js** file and add the following JavaScript:

        ````javascript
        // define controller
        var controllerId = 'shell';
        angular.module('app').controller(controllerId,
          ['$rootScope', '$route', 'dataContextAngular', 'common', 'config', shell]);
        // create controller
        function shell($rootScope, $route, dataContextAngular, common, config) {
          var vm = this;
          // Boolean property used to show/hide splash page
          vm.showSplashPage = true;
          // props to control the busy indicator
          vm.isBusy = true;
          // init controller
          init();
          function init() {
            // wire handler to successful route changes to
            //  - update the page title (for bookmarking)
            $rootScope.$on('$routeChangeSuccess', function (event, current, previous) {
              if (!$route.current || !$route.current.title) {
                $rootScope.pageTitle = '';
              } else {
                $rootScope.pageTitle = ' > ' + $route.current.title;
              }
            });
            // activate this controller, then initialize the app (making
            //   sure the lists in the SharePoint site are present)
            //   then hide the splash page
            common.activateController([dataContextAngular.initConfigurations()], controllerId)
              .then(function () {
                // app ready, hide splash
                vm.showSplashPage = false;
              })
              .catch(function (error) {
                // if receiving an error page redirect to the NoAuth (401) view
                if (error == 401) {
                  window.location = "/Home";
                }
              });
            // listen for the custom event when controllers are activated successfully
            $rootScope.$on(config.events.controllerActivateSuccess, function (data, args) {
              if (args && args.controllerId && args.controllerId != 'shell') {
                vm.isBusy = false;
              }
            });
          }
        }
        ````
        
        Save your changes and close the **App\layout\shell.js** file.

1. At this point you can implement a few views to view a list of all projects, view the details of a project and then create a project.
    1. First create the screen that shows a list of all projects. Open the **App\projects\list.html**.
    1. Locate the opening `<div>` and add the following **controller** directive to it:
    
        ````xml
        <div data-ng-controller="projectList as vm">
        ````

    1. Locate the first `<button>` and add the following **click** directive to it:
    
        ````xml
        <button type="button" class="btn btn-primary" data-ng-click="vm.goNewProject()">
        ````

    1. Locate the second `<button>` and add the following **click** directive to it:
    
        ````xml
        <button type="button" class="btn btn-default" data-ng-click="vm.goRefresh()">
        ````

    1. Scroll to the bottom of the file and look for the comment **project list**. Update the table row & cell to create a list of all projects and display their titles, but when someone clicks on the row, wire it up to an event handler on the controller:

        ````xml
        <tbody>
          <tr data-ng-repeat="proj in vm.projects" data-ng-click="vm.goProjectDetail(proj)">
            <td>{{proj.Title}}</td>
          </tr>
        </tbody>
        ````

    1. Finally, find the empty `<div>` that will be used for the modal dialog and add the **include** controller to include a modal dialog view:

        ````xml
        <div data-ng-include="'/App/projects/detailEdit.html'"></div>
        ````

        Save your changes and close the **App\projects\list.html** file.

    1. Now implement the controller. Open the **App\projects\list.js** file and add the following JavaScript to it:
        
        ````javascript
        // define controller
        var controllerId = 'projectList';
        angular.module('app').controller(controllerId,
          ['$location', '$timeout', '$modal', 'dataContextBreeze', 'common', projectList]);
        // create controller
        function projectList($location, $timeout, $modal, datacontext, common) {
          var vm = this;
          vm.goNewProject = goNewProject;
          vm.goRefresh = goRefresh;
          vm.goProjectDetail = goProjectDetail;
          // init controller
          init();
          function init() {
            common.activateController([getAllProjects()], controllerId);
          }
          // retrieve all projects form the server
          function getAllProjects(forceRefresh) {
            return datacontext.getProjects(forceRefresh)
              .then(function (data) {
                return vm.projects = data;
              })
              .catch(function (err) {
                throw new Error("error obtaining data: " + err);
              });
          }
          // refresh the project list
          function goRefresh() {
            // clear out the existing projects
            vm.projects = [];
            // force a refresh of all projects 
            return getAllProjects(true);
          }
          // redirects to the new project page
          function goNewProject() {
            var modal = $modal.open({
              templateUrl: 'myProjectModalContent.html',
              controller: 'projectDetailEdit',
              resolve: {
                project: function () { return datacontext.createProject(); }
              }
            });
            // update the list of projects with the local cache
            return modal.result.then(function() {
              return getAllProjects(false);
            });
          }
          // navigate to the detail page
          function goProjectDetail(project) {
            if (project && project.Id) {
              $location.path('/projects/' + project.Id);
            }
          }
        }
        ````

        Save your changes and close the **App\projects\list.js** file.

    1. Now implement the screen for creating a new, or editing an existing project. Open the **App\projects\detailEdit.html**.
    1. Locate the `<h2>` and replace it with the following to create a dynamic page title:
        
        ````xml
        <h2 data-ng-bind="pageTitle"></h2>
        ````

    1. Next, locate the `<input>` element on the page and add the following Angular **model** directive to implement two-way data binding with a model property:
        
        ````xml
        data-ng-model="project.Title"
        ````

    1. Finally, update the buttons to change their enabled/disabled state as well as their click handlers. Locate the `<div class="modal-footer">` and replace it with the following:
        
        ````xml
        <div class="modal-footer">
          <button type="button" class="btn btn-primary"
                  data-ng-disabled="!entityIsDirty"
                  data-ng-click="goSave()">save</button>
          <button type="button" class="btn btn-default"
                  data-ng-click="goCancel()">cancel</button>
          <button type="button" class="btn btn-danger"
                  data-ng-show="templateMode=='edit'"
                  data-ng-click="goDelete()">delete</button>
        </div>
        ````

        Save your changes and close the **App\projects\detailEdit.html** file.

    1. Now implement the controller. Open the **App\projects\detailEdit.js** file and add the following JavaScript to it:

        ````javascript
        // define controller
        var controllerId = 'projectDetailEdit';
        angular.module('app').controller(controllerId,
          ['$scope', '$modalInstance', '$location', 'dataContextBreeze', 'common', 'project', projectDetailEdit]);
        // create controller
        function projectDetailEdit($scope, $modalInstance, $location, datacontext, common, project) {
          var vm = $scope;
          vm.project = project;
          vm.goSave = goSave;
          vm.goCancel = goCancel;
          vm.goDelete = goDelete;
          // init controller
          init();
          function init() {
            if (project && +project.Id && project.Id > 0) {
              vm.pageTitle = "Edit Project";
              vm.templateMode = 'edit';
            } else {
              vm.pageTitle = "Create New Project";
              vm.templateMode = 'new';
            }
            // wire dirty checker
            wireUpEntityChangedHandler(vm.project);
          }
          // listen for changes to fields in the entity
          //  and update flag to enable/disable save button
          function wireUpEntityChangedHandler(entity) {
            vm.entityIsDirty = false;
            // when a property changes on the entity
            entity.entityAspect.propertyChanged.subscribe(
              function (args) {
                // if the title property changed, update flag
                if (args.propertyName == 'Title') {
                  vm.entityIsDirty = true;
                }
              });
          }
          // save the item
          function goSave() {
            return datacontext.saveChanges()
              .then(function () {
                $modalInstance.close(vm.project);
              });
          }
          // reset the changes
          function goCancel() {
            // cancel any changes to the item
            datacontext.revertChanges();
            // close the modal dialog
            $modalInstance.dismiss('cancel');
          }
          // delete entities
          function goDelete() {
            // get all references
            return datacontext.getProjectReferences(vm.project.Id)
              .then(function (references) {
                // loop through all references...
                for (var index = 0; index < references.length; index++) {
                  // set each to deleted, but supress the save call to the 
                  //  service... let the deletion of the project that will
                  //  follow be the one that issues the batch delete
                  datacontext.deleteEntity(references[index], true);
                }
              })
              .then(function () {
                return datacontext.deleteEntity(vm.project)
                  .then(function () {
                    // close the dialog & navigate back to the project list page
                    $modalInstance.close(
                      $location.path('/projects/')
                      );
                  });
              });
          }
        }
        ````

        Save your changes and close the **App\projects\detailEdit.js** file.

In this task you added code to existing JavaScript and HTML files to implement the single page application (SPA).


###Task 4 - Test the Application
Follow these steps to test the application.

1. When prompted, login to Azure AD using your Office 365 credentials.
1. On the homepage for the SPA, click the **PROJECT LIST** link.
1. This page display all projects that reside in the Project list in the SharePoint site. At first there won't be any projects...
1. Click the **new projects** link to open the **Create New Project** dialog.
1. Enter the name for a project and click **save**.
1. The view page should show the new project, as will the project list page.
1. Create a new more projects, then do a hard-refresh in your browser to force the SPA to clear it's local cache and retrieve the values from the SharePoint list again. This ensures data has been saved to the SharePoint list. You can also confirm this by navigating to the list in your SharePoint site.

Now you have completed testing the application.

##Summary
By completing this hands-on lab you learnt how to:
- Learn to configure Microsoft Azure AD for an Office 365 app
- Understand how and why you need to secure an OAuth access token
- Create a single page app
