/*
 * role: edit controller
 * desc: controller for the projects detail edit dialog
 */
(function () {
  'use strict';

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
})();