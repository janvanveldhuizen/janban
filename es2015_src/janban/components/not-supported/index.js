import angular from 'angular';

let notSupported = () => {
  return {
    template: require('./not-supported.html'),
    controller: 'NotSupportedCtrl',
    controllerAs: 'notSupported'
  }
};

class NotSupportedCtrl {
  constructor() {
    this.message = 'Sorry, this app can only be run as the home page of a folder in Outlook';
  }
}

const MODULE_NAME = 'notSupported';

angular.module(MODULE_NAME, [ ])
  .directive('notSupported', notSupported)
  .controller('NotSupportedCtrl', NotSupportedCtrl);

export default MODULE_NAME;