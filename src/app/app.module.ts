import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { AppComponent } from './app.component';
import { OutlookService } from './services/outlook.service';
import { NotSupportedComponent } from './not-supported/not-supported.component';

@NgModule({
  declarations: [
    AppComponent,
    NotSupportedComponent
  ],
  imports: [
    BrowserModule
  ],
  providers: [OutlookService],
  bootstrap: [AppComponent]
})
export class AppModule { }
