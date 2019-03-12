import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { AadHttpClient, HttpClientResponse} from '@microsoft/sp-http';

import {IK2Worklist, IK2Task} from './K2DataContracts';
import {K2DataReader} from './K2DataReader';
import {K2DataView} from './K2DataView';

import styles from './K2ApiWebPartWebPart.module.scss';
import * as strings from 'K2ApiWebPartWebPartStrings';

export interface IK2ApiWebPartWebPartProps {
  k2ServerURL: string;
}

export default class K2ApiWebPartWebPart extends BaseClientSideWebPart<IK2ApiWebPartWebPartProps> {
  private aadClient: AadHttpClient;
  private _dataReader: K2DataReader;

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error:any) => void): void => {
      this.context.aadHttpClientFactory
      .getClient('https://api.k2.com/')
      .then((client: AadHttpClient): void => {
        this.aadClient = client;
        resolve();
        this._dataReader = new K2DataReader(this.aadClient);
      }, err => reject(err));
    });

  }

  public render(): void {
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Loading...');
      this._dataReader.getData(this.properties.k2ServerURL)
      .then((worklist: IK2Worklist): void => {
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.domElement.innerHTML = K2DataView.getWorklistHtml(worklist);
      }, (err: any): void => {
        this.context.statusRenderer.renderError(this.domElement, err);
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.ConfigurationGroupName,
              groupFields: [
                PropertyPaneTextField('k2ServerURL', {
                  label: strings.k2ServerURLFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
