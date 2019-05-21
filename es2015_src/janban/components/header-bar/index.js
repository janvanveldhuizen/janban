import angular from 'angular';

let headerBar = () => {
  return {
    template: require('./header-bar.html'),
    controller: 'headerBarCtrl',
    controllerAs: 'headerBar'
  }
};

class headerBarCtrl {
  constructor() {
    this.message = 'this is the navbar';
  }
}

const MODULE_NAME = 'headerBar';

angular.module(MODULE_NAME, [ ])
  .directive('headerBar', headerBar)
  .controller('headerBarCtrl', headerBarCtrl);

export default MODULE_NAME;