import { Component, Input, Output, OnChanges, EventEmitter } from '@angular/core';

@Component({
  selector: 'app-header-bar',
  templateUrl: './header-bar.component.html',
  styleUrls: ['./header-bar.component.css']
})
export class HeaderBarComponent implements OnChanges {

  // privacyFilter: 
  // 0 = both
  // 1 = work
  // 2 = private
  @Input() privacyFilter: number; 
  @Output() privacyFilterChange = new EventEmitter<number>();
  
  constructor() {
    // this.privacyFilter = 2;
   }

   ngOnChanges() {
    // this.returnedString = 'number'+this.myStringArray[Number(this.inputNumber)];   
  }

  clickedRefresh = function() {
    alert('Refresh')
  }

  clickedReport = function() {
    alert('Report')
  }

  clickedConfig = function() {
    alert('Config')
  }

  clickedAbout = function() {
    alert('About')
  }

}
