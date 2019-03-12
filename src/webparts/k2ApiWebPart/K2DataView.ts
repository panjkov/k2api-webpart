import {IK2Worklist} from './K2DataContracts';

import styles from './K2ApiWebPartWebPart.module.scss';


export class K2DataView {

    public static getWorklistHtml(worklist: IK2Worklist) {
        var html = '';

        html = `
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
        <p>Available tasks: ${worklist.itemCount} in total</p>
       

</div>`;
        return html;
    }
}