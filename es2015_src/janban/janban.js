import angular from 'angular';

import '../assets/css/taskboard.css';
import 'bootstrap/dist/css/bootstrap.min.css';
import { notSupported } from './components/not-supported';
import { headerBar } from './components/header-bar';
import { outlookService } from './services/outlook-service.js';

let janban = () => {
  return {
    template: require('./janban.html'),
    controller: 'JanbanCtrl',
    controllerAs: 'janban'
  }
};

class JanbanCtrl {
  constructor() {
    this.isSupported = function() {
      alert('hi')
      try {
        let svc = new outlookService();
        alert(2);
          
      } catch (error) {
        alert(error)
      }
      return true;
    }
  }
  // get isSupported() {
  //   alert(1)
  //   // let service = new outlookService();
  //   // alert(service)
  //   // return service.isRunningInOutlook();
  // }
}

const MODULE_NAME = 'janban';

angular.module(MODULE_NAME, ['notSupported', 'headerBar'])
  .directive('janban', janban)
  .controller('JanbanCtrl', JanbanCtrl);

export default MODULE_NAME;