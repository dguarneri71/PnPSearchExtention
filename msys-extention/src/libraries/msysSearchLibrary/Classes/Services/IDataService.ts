import { ISearchResult } from "@pnp/sp/search";
import { FieldCollectionDataValue } from "../Entities/FieldCollectionDataValue";
import InvokeFlowResult from "../Entities/InvokeFlowResult";
import QueryData from "../Entities/QueryData";
import SettingItem from "../Entities/SettingItem";

export interface IDataService {
    getSearchResult(qquery: QueryData, count: number, moment: any): Promise<ISearchResult[]>;
    getSettingsBySpecificKey(listTitle: string, key: string): Promise<SettingItem[]>;
    getLabels(listTitle: string): Promise<FieldCollectionDataValue[]>;
    invokePowerAutomateFlowExtended(flowUrl: string, parameters: any, getData: boolean): Promise<InvokeFlowResult>;
}