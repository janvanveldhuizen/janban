import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { FontAwesomeModule } from '@fortawesome/angular-fontawesome';
import { library as FontLibrary } from '@fortawesome/fontawesome-svg-core';
import { faSyncAlt, faListAlt, faGlobeAmericas, faLock, faSuitcase, faWrench, faQuestion } from '@fortawesome/free-solid-svg-icons';

import { AppComponent } from './app.component';
import { OutlookService } from './services/outlook.service';
import { NotSupportedComponent } from './not-supported/not-supported.component';
import { HeaderBarComponent } from './header-bar/header-bar.component';

@NgModule({
  declarations: [
    AppComponent,
    NotSupportedComponent,
    HeaderBarComponent
  ],
  imports: [
    BrowserModule, 
    FontAwesomeModule
  ],
  providers: [OutlookService],
  bootstrap: [AppComponent]
})
export class AppModule { 
  constructor() {
    FontLibrary.add(faSyncAlt, faListAlt, faGlobeAmericas, faLock, faSuitcase, faWrench, faQuestion);
  }
}
