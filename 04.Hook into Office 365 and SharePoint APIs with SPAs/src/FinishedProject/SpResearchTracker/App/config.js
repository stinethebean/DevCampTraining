/*
 * role: global app configuration
 * desc: used for global app configuration
 */
(function () {
  'use strict';

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
})();