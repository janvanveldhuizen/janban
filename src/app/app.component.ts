import { Component } from '@angular/core';
import { OutlookService } from './services/outlook.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  canRun: boolean;
  privateOrWork: number;
  appMode: string;
  taskFilter: string;
    
  constructor(private outlookService: OutlookService){
    this.canRun = outlookService.isRunningInOutlook();
    this.privateOrWork = 2;
    this.appMode = 'kanban';
    this.taskFilter = 'qwerty';
  }

  taskFilterChanged(newFilter: string){
    this.taskFilter=newFilter;
  }

  titleClicked() {
    this.appMode="kanban";
  }

  refreshButtonClicked() {
    alert('refresh button clicked! now what?')
  }

  reportButtonClicked() {
    alert('report button clicked! now what?')
  }

  configButtonClicked() {
    this.appMode="config";
  }

  helpButtonClicked() {
    this.appMode="about";
  }
}
