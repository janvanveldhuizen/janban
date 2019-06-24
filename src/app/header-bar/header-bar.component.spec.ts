import { async, ComponentFixture, TestBed, tick } from '@angular/core/testing';
import { FontAwesomeModule } from '@fortawesome/angular-fontawesome';
import { FormsModule } from '@angular/forms';
import { library as FontLibrary } from '@fortawesome/fontawesome-svg-core';
import { faSyncAlt, faListAlt, faGlobeAmericas, faLock, faSuitcase, faWrench, faQuestion } from '@fortawesome/free-solid-svg-icons';

import { HeaderBarComponent } from './header-bar.component';


describe('HeaderBarComponent', () => {
  let component: HeaderBarComponent;
  let fixture: ComponentFixture<HeaderBarComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [HeaderBarComponent],
      imports: [FontAwesomeModule, FormsModule]
    })
      .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(HeaderBarComponent);
    FontLibrary.add(faSyncAlt, faListAlt, faGlobeAmericas, faLock, faSuitcase, faWrench, faQuestion);
    component = fixture.componentInstance;
    component.privacyFilter = 0;
    fixture.detectChanges();
    let dropdown = fixture.debugElement.query((de)=>{return de.nativeElement.id==="privacy-select"});
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });

  it('should have dropdown selected 0 when input privacyFilter = 0', () => {
    component.privacyFilter = 0;
    fixture.detectChanges();
    expect(component).toBeTruthy();
    const compiled = fixture.debugElement.nativeElement;
    const dropdown = compiled.querySelector('#privacy-select');
    console.log(dropdown.selectedOptions)
    expect(dropdown.selectedOptions(0).value).toEqual('0');
  });


  it('should have dropdown selected 1 when input privacyFilter = 1', () => {
    component.privacyFilter = 1;
    fixture.detectChanges();
    expect(component).toBeTruthy();
    const compiled = fixture.debugElement.nativeElement;
    const dropdown = compiled.querySelector('#privacy-select');
    expect(dropdown.options[dropdown.selectedIndex].value).toEqual('1');
  });


  it('should have dropdown selected 2 when input privacyFilter = 2', () => {
    component.privacyFilter = 2;
    fixture.detectChanges();
    expect(component).toBeTruthy();
    const compiled = fixture.debugElement.nativeElement;
    const dropdown = compiled.querySelector('#privacy-select');
    expect(dropdown.options[dropdown.selectedIndex].value).toEqual('2');
  });
});
