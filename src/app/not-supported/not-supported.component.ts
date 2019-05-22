import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-not-supported',
  templateUrl: './not-supported.component.html',
  styleUrls: ['./not-supported.component.css']
})
export class NotSupportedComponent implements OnInit {
  message: string;

  constructor() { 
     this.message = 'Sorry, this app can only be run as the home page of a folder in Outlook';
  }

  ngOnInit() {
  }

}
