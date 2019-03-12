
export interface IK2Worklist {
    itemCount: string;
    tasks: IK2Task[];
}

export interface IK2Task {
    serialNumber: string;
    formURL: string;
    viewFlowURL: string;
    workflowDisplayName: string;
    workflowInstanceID: number;
    workflowInstanceFolio: string;
    activityName: string;

}