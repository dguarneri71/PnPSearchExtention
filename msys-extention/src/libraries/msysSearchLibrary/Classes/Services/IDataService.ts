import { ISearchResult } from "@pnp/sp/search";
import InvokeFlowResult from "../Entities/InvokeFlowResult";
import QueryData from "../Entities/QueryData";
import SettingItem from "../Entities/SettingItem";

export interface IDataService {
    getSearchResult(qquery: QueryData, count: number): Promise<ISearchResult[]>;
    getSettingsBySpecificKey(listTitle: string, key: string): Promise<SettingItem[]>;
    invokePowerAutomateFlowExtended(flowUrl: string, parameters: any, getData: boolean): Promise<InvokeFlowResult>;
}