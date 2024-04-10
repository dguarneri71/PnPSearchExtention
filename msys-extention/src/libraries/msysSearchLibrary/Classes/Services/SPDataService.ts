import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IDataService } from "./IDataService";
import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';
// import { SPUser } from "@microsoft/sp-page-context";
import { stringIsNullOrEmpty } from "@pnp/common";
// import { Guid } from '@microsoft/sp-core-library';
import "@pnp/sp/search";
import { ISearchQuery, SearchResults, ISearchResult } from "@pnp/sp/search";
import { IList, IViewInfo, sp, Web } from "@pnp/sp/presets/all";
import { PageContext } from '@microsoft/sp-page-context';
import QueryData from '../Entities/QueryData';
import InvokeFlowResult from "../Entities/InvokeFlowResult";
import { FieldNames } from '../Constants';
import SettingItem from '../Entities/SettingItem';
import { FieldCollectionDataValue } from '../Entities/FieldCollectionDataValue';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { Helper } from '../../Helpers/Helper';
import DownloadFile from '../Entities/DownloadFile';
import * as JSZip from 'jszip';
import { saveAs } from 'file-saver';

const LOG_SOURCE: string = 'SPDataService';

export default class SPDataService implements IDataService {
    private _httpClient: HttpClient;

    constructor(context: PageContext, httpClient: HttpClient) {
        console.log(LOG_SOURCE + " - constructor - absoluteUrl: ", context.web.absoluteUrl);
        this._httpClient = httpClient;
        sp.setup({
            pageContext: context
        });
    }

    /**
     * Esegue una query al motore di ricerca di SharePoint, tenendo conto anche dei RefinementFilters
     * NB TODO: verificare bene perchè non estrae correttamente tutti i valori - confrontare i parametri delle query iniziale con quelli ricostruiti
     * @param query 
     * @param moment
     * @param progressCallback 
     * @returns 
     */
    public async getSearchResult(query: QueryData, moment: any, progressCallback: (percentComplete: number, partial: number, total: number) => void): Promise<ISearchResult[]> {
        console.log(LOG_SOURCE + " - getSearchResult() - query: ", query);
        console.log(LOG_SOURCE + " - getSearchResult() - total row count: ", query.totalItemsCount);
        const itemPerPage = query.itemsCountPerPage;
        console.log(LOG_SOURCE + " - getSearchResult() - items per page: ", itemPerPage);
        const totalPageNum = Math.ceil(query.totalItemsCount / itemPerPage);
        console.log(LOG_SOURCE + " - getSearchResult() - num of page: ", totalPageNum);
        const rowLimit = totalPageNum > 1 ? itemPerPage : query.totalItemsCount;
        console.log(LOG_SOURCE + " - getSearchResult() - num of page: ", totalPageNum);
        //Recupero i risultati paginando
        let results: ISearchResult[] = [];
        let partialRowCount: number = 0;
        let errMsg: string = "";

        try {
            let searchQuery: ISearchQuery = {};
            searchQuery.Querytext = query.queryText;
            searchQuery.RowLimit = rowLimit;
            searchQuery.SelectProperties = query.SelectProperties;
            searchQuery.EnableQueryRules = query.enableQueryRules;
            searchQuery.SourceId = query.resultSourceId;
            searchQuery.QueryTemplate = query.queryTemplate.replace("{verticals.value}", query.verticalValue);
            searchQuery.TrimDuplicates = query.trimDuplicates;
            searchQuery.SortList = query.sortList;

            let refinementFilters: string[] = !isEmpty(query.refinementFilters) ? [query.refinementFilters] : [];

            if (!isEmpty(query.selectedFilters)) {
                // Set list of refiners to retrieve
                searchQuery.Refiners = query.filtersConfiguration.map(filterConfig => {
                    // Special case with Date managed properties
                    const regexExpr = "(RefinableDate\\d+)(?=,|$)|" +
                        "(RefinableDateInvariant00\\d+)(?=,|$)|" +
                        "(RefinableDateSingle\\d+)(?=,|$)|" +
                        "(LastModifiedTime)(?=,|$)|" +
                        "(LastModifiedTimeForRetention)(?=,|$)|" +
                        "(Created)(?=,|$)|" +
                        "(Date\\d+)(?=,|$)|" +
                        "(EndDate)(?=,|$)|" +
                        "(.+OWSDATE)(?=,|$)|" +
                        "(EventsRollUpEndDate)(?=,|$)|" +
                        "(EventsRollUpStartDate)(?=,|$)|" +
                        "(FirstPublishedDate)(?=,|$)|" +
                        "(ImageDateCreated)(?=,|$)|" +
                        "(LastAnalyticsUpdateTime)(?=,|$)|" +
                        "(ModifierDates)(?=,|$)|" +
                        "(ClassificationLastScan)(?=,|$)|" +
                        "(ComplianceTagWrittenTime)(?=,|$)|" +
                        "(ContentModifiedTime)(?=,|$)|" +
                        "(DocumentAnalyticsLastActivityTimestamp)(?=,|$)|" +
                        "(ExpirationTime)(?=,|$)|" +
                        "(LastSharedByTime)(?=,|$)|" +
                        "(StartDate)(?=,|$)|" +
                        "(TagEventDate)(?=,|$)|" +
                        "(processingtime)(?=,|$)|" +
                        "(ExtractedDate)(?=,|$)";

                    const refinableDateRegex = new RegExp(regexExpr.replace(/\s+/gi, ''), 'gi');
                    if (refinableDateRegex.test(filterConfig.filterName)) {
                        const pastYear = moment(new Date()).subtract(1, 'years').subtract('minutes', 1).toISOString();
                        const past3Months = moment(new Date()).subtract(3, 'months').subtract('minutes', 1).toISOString();
                        const pastMonth = moment(new Date()).subtract(1, 'months').subtract('minutes', 1).toISOString();
                        const pastWeek = moment(new Date()).subtract(1, 'week').subtract('minutes', 1).toISOString();
                        const past24hours = moment(new Date()).subtract(24, 'hours').subtract('minutes', 1).toISOString();
                        const today = new Date().toISOString();

                        return `${filterConfig.filterName}(discretize=manual/${pastYear}/${past3Months}/${pastMonth}/${pastWeek}/${past24hours}/${today})`;
                    }
                    else {
                        return filterConfig.filterName;
                    }
                }).join(',');

                // Get refinement filters
                if (query.selectedFilters.length > 0) {
                    // Make sure, if we have multiple filters, at least two filters have values to avoid apply an operator ('or','and') on only one condition failing the query.
                    if (query.selectedFilters.length > 1 && query.selectedFilters.filter(selectedFilter => selectedFilter.values.length > 0).length > 1) {
                        const refinementString = Helper.buildFqlRefinementString(query.selectedFilters, moment).join(',');
                        if (!isEmpty(refinementString)) {
                            refinementFilters = refinementFilters.concat([`${query.filterOperator}(${refinementString})`]);
                        }

                    } else {
                        refinementFilters = refinementFilters.concat(Helper.buildFqlRefinementString(query.selectedFilters, moment));
                    }
                }
            }

            searchQuery.RefinementFilters = refinementFilters;
            console.log(LOG_SOURCE + " - getSearchResult() - searchQuery: ", searchQuery);

            //#region PRIMO METODO DI PAGINAZIONE
            const searchResults: SearchResults = await sp.search(searchQuery);
            console.log(LOG_SOURCE + " - getSearchResult() - searchResults: ", searchResults);
            results = searchResults.PrimarySearchResults ? searchResults.PrimarySearchResults : [];
            partialRowCount = searchResults.RowCount;
            progressCallback(partialRowCount / query.totalItemsCount, partialRowCount, query.totalItemsCount);

            if (totalPageNum > 1) {
                for (let page = 2; page <= totalPageNum; page++) {
                    console.log(LOG_SOURCE + " - getSearchResult() - Page: ", page);
                    let size = (page * itemPerPage) < query.totalItemsCount ? itemPerPage : query.totalItemsCount - partialRowCount;
                    console.log(LOG_SOURCE + " - getSearchResult() - page size: ", size);
                    //let itSearchResuls: SearchResults = await searchResults.getPage(page, size);
                    let itSearchResuls: SearchResults = await searchResults.getPage(page);
                    console.log(LOG_SOURCE + " - getSearchResult() - itSearchResuls: ", itSearchResuls);
                    if (itSearchResuls) {
                        results = results.concat(itSearchResuls.PrimarySearchResults);
                        partialRowCount = partialRowCount + itSearchResuls.RowCount;
                        progressCallback(partialRowCount / query.totalItemsCount, partialRowCount, query.totalItemsCount);
                    }
                }
            }


            //#endregion

            //#region SECONDO METODO DI PAGINAZIONE - da provare
            /*
            for (let page = 1; page <= totalPageNum; page++) {
                console.log(LOG_SOURCE + " - getSearchResult() - Page: ", page);
                searchQuery.StartRow = itemPerPage * (page - 1) + 1;
                console.log(LOG_SOURCE + " - getSearchResult() - StartRow: ", searchQuery.StartRow);
                let searchResults: SearchResults = await sp.search(searchQuery);
                console.log(LOG_SOURCE + " - getSearchResult() - searchResults: ", searchResults);
                results = results.concat(searchResults.PrimarySearchResults);
                partialRowCount = partialRowCount + searchResults.RowCount;
                progressCallback(partialRowCount / query.totalItemsCount, partialRowCount, query.totalItemsCount);
            }
            */
            //#endregion

            console.log(LOG_SOURCE + " - getSearchResult() - results: ", results);
        } catch (e) {
            // are we dealing with an HttpRequestError?
            if (e?.isHttpRequestError) {

                // we can read the json from the response
                const json = await e.response.json();

                // if we have a value property we can show it
                errMsg = typeof json["odata.error"] === "object" ? json["odata.error"].message.value : e.message;

            } else {
                // not an HttpRequestError so we just log message
                errMsg = e.message;
            }
        }

        return new Promise<ISearchResult[]>((resolve, reject) => {
            if (stringIsNullOrEmpty(errMsg)) {
                resolve(results);
            }
            else {
                reject(errMsg);
            }
        });
    }

    /**
     * Invoca un flow Power Automate, che ha un HTTPTrigger
     * @param flowUrl La URL del flow Power Automate
     * @param parameters I parametri del flow
     * @param getData Flag per indicare se c'è una risposta da parte del flow da memorizzare
     * @returns 
     */
    public invokePowerAutomateFlowExtended(flowUrl: string, parameters: any, getData: boolean): Promise<InvokeFlowResult> {
        const postURL: string = flowUrl;

        const body: string = JSON.stringify(parameters);

        const requestHeaders: Headers = new Headers();
        requestHeaders.append('Content-type', 'application/json');

        const httpClientOptions: IHttpClientOptions = {
            body: body,
            headers: requestHeaders
        };

        console.log(LOG_SOURCE + " - invokePowerAutomateFlowExtended");
        console.log(LOG_SOURCE + " - invokePowerAutomateFlowExtended url and parameters: ", flowUrl, parameters);

        return new Promise<InvokeFlowResult>(res => {
            let result: InvokeFlowResult = new InvokeFlowResult();
            this._httpClient.post(
                postURL,
                HttpClient.configurations.v1,
                httpClientOptions)
                .then((response: HttpClientResponse) => {
                    console.log(LOG_SOURCE + " - invokePowerAutomateFlowExtended Response: ", response);
                    if (!response.ok) {
                        console.log(LOG_SOURCE + " - invokePowerAutomateFlowExtended FALSE");
                        result.success = false;
                        res(result);
                    }
                    else {
                        console.log(LOG_SOURCE + " - invokePowerAutomateFlowExtended TRUE");
                        if (getData) {
                            console.log(LOG_SOURCE + " - invokePowerAutomateFlowExtended get flow response");
                            response.json().then(data => {
                                console.log(LOG_SOURCE + " - invokePowerAutomateFlowExtended get flow response data: ", data);
                                result.success = true;
                                result.data = data;
                                res(result);
                            }).catch(error => { //questo caso non so se può accadere
                                console.log(LOG_SOURCE + " - invokePowerAutomateFlowExtended get flow response error: ", error);
                                result.success = false;
                                result.error = error;
                                res(result);
                            });
                        }
                        else {
                            result.success = true;
                            result.data = null;
                            res(result);
                        }
                    }
                }).catch(error => {
                    console.log(LOG_SOURCE + " - invokePowerAutomateFlowExtended FALSE: ", error);
                    result.success = false;
                    result.error = error;
                    res(result);
                });
        });
    }

    /**
     * 
     * @param listTitle 
     */
    public getLabels(listTitle: string): Promise<FieldCollectionDataValue[]> {
        console.log(LOG_SOURCE + " - getLabels() - listTitle: ", listTitle);
        let results: FieldCollectionDataValue[] = [];
        return new Promise<FieldCollectionDataValue[]>(async (res, reject) => {
            let labelsList: IList = sp.web.lists.getByTitle(listTitle);

            const view: IViewInfo = await labelsList.defaultView();
            console.log(LOG_SOURCE + " - getLabels() - default view: ", view);
            console.log(LOG_SOURCE + " - getLabels() - default view query: ", view.ViewQuery);

            const xml = '<View><Query>' + view.ViewQuery + '</Query></View>';

            //Recupero gli item in base alla query della vista di default
            labelsList.getItemsByCAMLQuery({'ViewXml' : xml}).then(items => {
                console.log(LOG_SOURCE + " - getLabels() - items: ", items);
                for (let index = 0; index < items.length; index++) {
                    const element = items[index];
                    let result: FieldCollectionDataValue = new FieldCollectionDataValue();
                    result.Title = element.Title;
                    result.DisplayName = element["Label"];
                    result.Type = element["ColumnType"];
                    result.Format = element["Format"];
                    //result.Order = element["Order0"];
                    results.push(result);
                }
                console.log(LOG_SOURCE + " - getLabels() - labels: ", results);
                res(results);
            }).catch(reason => {
                Object.keys(reason).forEach(prop => console.log(LOG_SOURCE + " - prop: ", prop, reason[prop]));
                reject(reason);
            });

            //Vecchio metodo
            /* labelsList.items.select("ID", "Title", "Label", "ColumnType", "Format", "Order0").orderBy("Order0").get().then(items => {
                console.log(LOG_SOURCE + " - getLabels() - items: ", items);
                for (let index = 0; index < items.length; index++) {
                    const element = items[index];
                    let result: FieldCollectionDataValue = new FieldCollectionDataValue();
                    result.Title = element.Title;
                    result.DisplayName = element["Label"];
                    result.Type = element["ColumnType"];
                    result.Format = element["Format"];
                    result.Order = element["Order0"];
                    results.push(result);
                }
                console.log(LOG_SOURCE + " - getLabels() - labels: ", results);
                res(results);
            }).catch(reason => {
                Object.keys(reason).forEach(prop => console.log(LOG_SOURCE + " - prop: ", prop, reason[prop]));
                reject(reason);
            }); */
        });
    }

    /**
     * Recupera tutti i record nella lista di configurazione con la chiave passata
     * @param listTitle Titolo della lista di configurazione
     * @param key Chiave di ricerca delle configurazioni
     * @returns 
     */
    public getSettingsBySpecificKey(listTitle: string, key: string): Promise<SettingItem[]> {
        return new Promise<any[]>((res, reject) => {
            this.getSettings(listTitle, key).then(items => {
                let results: SettingItem[] = [];
                for (let index = 0; index < items.length; index++) {
                    const item = items[index];
                    let result = new SettingItem(item[FieldNames.SettingTitle], item[FieldNames.SettingValue]);
                    results.push(result);
                }
                res(results);
            }).catch(reason => {
                reject(reason);
            });
        });
    }

    // https://stuk.github.io/jszip/
    // funziona solo sul sito corrente
    public async downloadZipFile(files: DownloadFile[], progressCallback: any): Promise<void> {
        var zip = new JSZip();

        for (let index = 0; index < files.length; index++) {
            const element = files[index];
            try {
                console.log(LOG_SOURCE + " - saveFile() - file: ", element);
                progressCallback(files, index, "", false);
                const file = sp.web.getFileByServerRelativePath(element.serverRelativeUrl);
                let content: ArrayBuffer = await file.getBuffer();
                console.log(LOG_SOURCE + " - saveFile() - blob: ", content);
                zip.file(element.filename, content);
            } catch (reason) {
                console.log(LOG_SOURCE + " - saveFile() - error: ", reason);
                progressCallback(files, index, reason, false);
            }
        }

        console.log(LOG_SOURCE + " - saveFile() - zip: ", zip);
        zip.generateAsync({ type: "blob" })
            .then((content) => {
                console.log(LOG_SOURCE + " - saveFile() - save zip");
                // see FileSaver.js
                saveAs(content, "dowbload.zip");
                progressCallback(files, files.length, "Finished", true);
            });
    }

    /******** Private Methods  ********/
    private getSettings(listTitle: string, key: string): Promise<any[]> {
        let _filter = "Key eq '" + key + "'";
        console.log(LOG_SOURCE + " - getSettings() - filter: ", _filter);
        console.log(LOG_SOURCE + " - getSettings() - listTitle: ", listTitle);
        return new Promise<any[]>((res, reject) => {
            let configList: IList = sp.web.lists.getByTitle(listTitle); //sp.web.getList(listSettingsUrl);
            configList.items.filter(_filter).get().then(items => {
                console.log(LOG_SOURCE + " - getSettings() - items: ", items);
                res(items);
            }).catch(reason => {
                reject(reason);
            });
        });
    }
}
