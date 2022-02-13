export default class QueryData {
    public queryText: string;
    public enableQueryRules: boolean;
    public queryTemplate: string;
    public resultSourceId: string;

    constructor(queryText:string, enableQueryRules: boolean, queryTemplate: string, resultSourceId: string) { 
        this.enableQueryRules = enableQueryRules;
        this.queryTemplate = queryTemplate;
        this.queryText = queryText;
        this.resultSourceId = resultSourceId;
    }
}