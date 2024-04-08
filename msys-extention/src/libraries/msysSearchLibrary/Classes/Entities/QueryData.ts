import { IDataFilter, IDataFilterConfiguration } from "@pnp/modern-search-extensibility";
import { ISort, SortDirection } from "@pnp/sp/search";

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
    public verticalValue: string;
    public itemsCountPerPage: number;
    public trimDuplicates: boolean;
    public totalItemsCount: number;
    public sortList : ISort[] = [];

    constructor(queryProps: any, selectedProperties: string[]) {

        let queryText: string = "*";
        if (queryProps.content["inputQueryText"]) {
            queryText = queryProps.content["inputQueryText"];
        }
        let totalItemsCount: number = queryProps.content["data"]["totalItemsCount"];
        let dataSourceProperties = queryProps.content["properties"]["dataSourceProperties"];
        let enableQueryRules: boolean = dataSourceProperties["enableQueryRules"];
        let queryTemplate: string = dataSourceProperties["queryTemplate"];
        let resultSourceId: string = dataSourceProperties["resultSourceId"];
        let useVertical: boolean = queryProps.content["properties"]["useVerticals"];
        console.log("QueryData - useVertical: ", useVertical);
        let selectedFilters: IDataFilter[] = queryProps.content["filters"]["selectedFilters"] as IDataFilter[];
        console.log("QueryData - selectedFilters: ", selectedFilters);
        let filtersConfiguration: IDataFilterConfiguration[] = queryProps.content["filters"]["filtersConfiguration"] as IDataFilterConfiguration[];
        console.log("QueryData - filtersConfiguration: ", filtersConfiguration);
        let filterOperator: string = queryProps.content["filters"]["filterOperator"];
        console.log("QueryData - filterOperator: ", filterOperator);
        let refinementFilters: string = dataSourceProperties["refinementFilters"];
        console.log("QueryData - refinementFilters: ", refinementFilters);
        let itemsCountPerPage: number = queryProps.content["properties"]["paging"]["itemsCountPerPage"];
        console.log("QueryData - itemsCountPerPage: ", itemsCountPerPage);
        let trimDuplicates: boolean = dataSourceProperties["trimDuplicates"];
        console.log("QueryData - trimDuplicates: ", trimDuplicates);
        let _sortList: any[] = dataSourceProperties["sortList"] as any[];
        console.log("QueryData - sortList: ", _sortList);

        let verticalValue: string = null;
        if (useVertical) {
            verticalValue = queryProps.content["verticals"]["selectedVertical"]["value"];
            console.log("QueryData - verticalValue: ", verticalValue);
        }

        this.enableQueryRules = enableQueryRules;
        this.queryTemplate = queryTemplate;
        this.queryText = queryText;
        this.resultSourceId = resultSourceId;
        this.SelectProperties = selectedProperties;
        this.selectedFilters = selectedFilters;
        this.refinementFilters = refinementFilters;
        this.filtersConfiguration = filtersConfiguration;
        this.filterOperator = filterOperator;
        this.verticalValue = verticalValue;
        this.itemsCountPerPage = itemsCountPerPage;
        this.totalItemsCount = totalItemsCount;
        this.trimDuplicates = trimDuplicates;

        _sortList.forEach(element => {
            let sortField: string = element["sortField"];
            let sortDirection: SortDirection = element["sortDirection"] ? element["sortDirection"] : SortDirection.Ascending;
            this.sortList.push({Property: sortField, Direction: sortDirection});
        });
    }
}