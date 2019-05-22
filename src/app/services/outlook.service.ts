import { Injectable } from '@angular/core';
import { environment } from 'src/environments/environment';

@Injectable()
export class OutlookService {
    constructor() {
    }

    isRunningInOutlook(): boolean {
        if (environment.production) {
        // use eval() here to cheat the Typescript transpiler
            return (window.external !== undefined && eval('window.external.OutlookApplication') !== undefined);
        } else {
            return true;
        }
    }
}