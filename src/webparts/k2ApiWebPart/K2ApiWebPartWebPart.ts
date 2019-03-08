import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { AadHttpClient, HttpClientResponse} from '@microsoft/sp-http';

import styles from './K2ApiWebPartWebPart.module.scss';
import * as strings from 'K2ApiWebPartWebPartStrings';
import { tasks } from '@microsoft/teams-js';

export interface IK2ApiWebPartWebPartProps {
  description: string;
  k2ServerURL: string;
}

export default class K2ApiWebPartWebPart extends BaseClientSideWebPart<IK2ApiWebPartWebPartProps> {
  private aadClient: AadHttpClient;

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error:any) => void): void => {
      this.context.aadHttpClientFactory
      .getClient('https://api.k2.com/')
      .then((client: AadHttpClient): void => {
        this.aadClient = client;
        resolve();
      }, err => reject(err));
    });

  }

  public render(): void {
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'k2ApiWebPart');

    this.aadClient.get(this.properties.k2ServerURL + '/api/workflow/preview/tasks', AadHttpClient.configurations.v1)
      .then((res: HttpClientResponse): Promise<any> => {
        
        return res.json();
      })
      .then((worklist: any): void => {
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.domElement.innerHTML = `
      <div class="${ styles.k2ApiWebPart }">

      <div class="${ styles.row }">
      <div class="${ styles.column }">
        <span class="${ styles.title }">K2 Worklist</span>
        </div>
        </div>
              <table>
              <tr class="ms-Grid-row"><th class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-bgColor-themeLight  ms-font-m-plus">Form URL</th><th class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg3 ms-bgColor-themeLight  ms-font-m-plus">Workflow Name</th><th class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg3 ms-bgColor-themeLight  ms-font-m-plus">Folio</th><th class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-bgColor-themeLight  ms-font-m-plus">Activity</th><th class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-bgColor-themeLight  ms-font-m-plus">Viewflow</th></tr>
              ${worklist.tasks.map(t => `<tr class="ms-Grid-row"><td class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m"><a href="${t.formURL}" target="_blank" class="${ styles.button }">Open Form</a></td><td class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg3 ms-font-m">${t.workflowDisplayName}</td><td class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg3 ms-font-m">${t.workflowInstanceFolio}</td><td class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">${t.activityName}</td><td class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m"><a href="${t.viewFlowURL}" target="_blank">View</a></td>`).join('')}
              </table>
              <p>Available tasks: ${worklist.itemCount}</p>
             

      </div>`;
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
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
