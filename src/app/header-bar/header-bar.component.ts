import { Component, Input, Output, DoCheck, EventEmitter } from '@angular/core';
import { text } from '@fortawesome/fontawesome-svg-core';

@Component({
  selector: 'app-header-bar',
  templateUrl: './header-bar.component.html',
  styleUrls: ['./header-bar.component.css'],
})
export class HeaderBarComponent implements DoCheck {
  previousFilter: string;
  doCheckCounter: number;
  // privacyFilter: 
  // 0 = both
  // 1 = work
  // 2 = private
  @Input() privacyFilter: number;
  @Input() textFilter: string;
  @Output() privacyFilterChange = new EventEmitter<number>();
  @Output() titleClicked = new EventEmitter();
  @Output() refreshButtonClicked = new EventEmitter();
  @Output() reportButtonClicked = new EventEmitter();
  @Output() configButtonClicked = new EventEmitter();
  @Output() helpButtonClicked = new EventEmitter();
  @Output() textFilterChanged = new EventEmitter<string>();

  constructor() {
    this.previousFilter = this.textFilter;
    this.doCheckCounter = 0;
  }

  ngDoCheck() {
    if (this.textFilter !== this.previousFilter){
      // alert('previousFilter=' + this.previousFilter + ' textFilter=' + this.textFilter)
      this.previousFilter = this.textFilter;
      this.textFilterChanged.emit(this.textFilter);
      this.doCheckCounter++;
    }
  }

  clickedTitle = function () {
    this.titleClicked.emit();
  }

  clickedRefresh = function () {
    this.refreshButtonClicked.emit();
  }

  clickedReport = function () {
    this.reportButtonClicked.emit();
  }

  clickedConfig = function () {
    this.configButtonClicked.emit();
  }

  clickedAbout = function () {
    this.helpButtonClicked.emit();
  }

}
