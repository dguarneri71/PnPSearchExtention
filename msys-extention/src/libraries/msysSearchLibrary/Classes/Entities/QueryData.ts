import { IDataFilter, IDataFilterConfiguration } from "@pnp/modern-search-extensibility";

export default class QueryData {
    public queryText: string;
    public enableQueryRules: boolean;
    public queryTemplate: string;
    public resultSourceId: string;
    public SelectProperties: string[];
    public selectedFilters: IDataFilter[];
    public refinementFilters: string;
    public filtersConfiguration: IDataFilterConfiguration[];
    public filterOperator: string;

    constructor(queryText: string, enableQueryRules: boolean, queryTemplate: string, resultSourceId: string,
        selectedProperties: string[], filtersConfiguration: IDataFilterConfiguration[],
        selectedFilters: IDataFilter[], refinementFilters: string, filterOperator: string) {
        this.enableQueryRules = enableQueryRules;
        this.queryTemplate = queryTemplate;
        this.queryText = queryText;
        this.resultSourceId = resultSourceId;
        this.SelectProperties = selectedProperties;
        this.selectedFilters = selectedFilters;
        this.refinementFilters = refinementFilters;
        this.filtersConfiguration = filtersConfiguration;
        this.filterOperator = filterOperator;
    }
}