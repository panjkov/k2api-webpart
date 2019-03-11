import { IWebPartContext } from '@microsoft/sp-webpart-base';

import { AadHttpClient, HttpClientResponse} from '@microsoft/sp-http';

import {IK2Worklist} from './K2DataContracts';

export class K2DataReader {
    private  _aadClient: AadHttpClient;

    constructor(aadClient: AadHttpClient) {
       this._aadClient = aadClient;
    }

    public getData(k2ServerURL): Promise<IK2Worklist> {
        return this._aadClient.get(k2ServerURL + '/api/workflow/preview/tasks', AadHttpClient.configurations.v1)
        .then((res: HttpClientResponse): Promise<any> => {          
          return res.json();
        });
    }
}

