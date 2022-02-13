import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IDataService } from "./IDataService";
import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';
// import { SPUser } from "@microsoft/sp-page-context";
// import { stringIsNullOrEmpty } from "@pnp/common";
// import { Guid } from '@microsoft/sp-core-library';
import "@pnp/sp/search";
import { ISearchQuery, SearchResults, ISearchResult } from "@pnp/sp/search";
import { IList, sp, Web } from "@pnp/sp/presets/all";
import { PageContext } from '@microsoft/sp-page-context';
import QueryData from '../Entities/QueryData';
import InvokeFlowResult from "../Entities/InvokeFlowResult";
import { FieldNames } from '../Constants';
import SettingItem from '../Entities/SettingItem';

const LOG_SOURCE: string = 'SPDataService';

export default class SPDataService implements IDataService {
    private _httpClient: HttpClient;
    //private _context: PageContext;

    /* constructor(absoluteUrl: string) {
        console.log(LOG_SOURCE + " - constructor - absoluteUrl: ", absoluteUrl);
        sp.setup({
            pageContext: {
                web: {
                    absoluteUrl: absoluteUrl
                }
            }
        });
    } */

    constructor(context: PageContext, httpClient: HttpClient) {
        console.log(LOG_SOURCE + " - constructor - absoluteUrl: ", context.web.absoluteUrl);
        this._httpClient = httpClient;
        sp.setup({
            pageContext: context
        });
    }

    public async getSearchResult(query: QueryData, count: number): Promise<ISearchResult[]> {
        console.log(LOG_SOURCE + " - getSearchResult(): ", query, count);

        const searchResults: SearchResults = await sp.search(<ISearchQuery>{
            Querytext: query.queryText,
            RowLimit: count,
            SelectProperties: ["Filename", "FileType", "FileExtension", "Path"],
            EnableQueryRules: query.enableQueryRules,
            SourceId: query.resultSourceId,
            QueryTemplate: query.queryTemplate,
            TrimDuplicates: false
        });

        console.log(LOG_SOURCE + " - getSearchResult() - ElapsedTime: ", searchResults.ElapsedTime);
        console.log(LOG_SOURCE + " - getSearchResult() - RowCount: ", searchResults.RowCount);
        console.log(LOG_SOURCE + " - getSearchResult() - PrimarySearchResults: ", searchResults.PrimarySearchResults);

        return new Promise<ISearchResult[]>(res => {
            res(searchResults.PrimarySearchResults);
        });
    }

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