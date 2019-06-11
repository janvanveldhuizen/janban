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
    
  constructor(private outlookService: OutlookService){
    this.canRun = outlookService.isRunningInOutlook();
    this.privateOrWork = 2;
  }
}
