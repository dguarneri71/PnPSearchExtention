export default class InvokeFlowResult { 
    public success: boolean;
    public data: any;
    public error: any;

    constructor() {
        this.data = null;
        this.error = null;
    }
}