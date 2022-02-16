export default class QueryData {
    public queryText: string;
    public enableQueryRules: boolean;
    public queryTemplate: string;
    public resultSourceId: string;
    public SelectProperties: string[];

    constructor(queryText:string, enableQueryRules: boolean, queryTemplate: string, resultSourceId: string, selectedProperties) { 
        this.enableQueryRules = enableQueryRules;
        this.queryTemplate = queryTemplate;
        this.queryText = queryText;
        this.resultSourceId = resultSourceId;
        this.SelectProperties = selectedProperties;
    }
}