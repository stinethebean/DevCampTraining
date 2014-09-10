/*
 * role: view controller
 * desc: controller for the shell view
 */
(function () {
  'use strict';

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
})();