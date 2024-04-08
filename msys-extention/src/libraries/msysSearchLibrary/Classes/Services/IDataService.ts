import { ISearchResult } from "@pnp/sp/search";
import DownloadFile from "../Entities/DownloadFile";
import { FieldCollectionDataValue } from "../Entities/FieldCollectionDataValue";
import InvokeFlowResult from "../Entities/InvokeFlowResult";
import QueryData from "../Entities/QueryData";
import SettingItem from "../Entities/SettingItem";

export interface IDataService {
    //getSearchResult(query: QueryData, totalItemsCount: number, moment: any, progressCallback: (percentComplete: number, partial: number, total: number) => void): Promise<ISearchResult[]>;
    getSearchResult(query: QueryData, moment: any, progressCallback: (percentComplete: number, partial: number, total: number) => void): Promise<ISearchResult[]>;
    getSettingsBySpecificKey(listTitle: string, key: string): Promise<SettingItem[]>;
    getLabels(listTitle: string): Promise<FieldCollectionDataValue[]>;
    invokePowerAutomateFlowExtended(flowUrl: string, parameters: any, getData: boolean): Promise<InvokeFlowResult>;
    downloadZipFile(files: DownloadFile[], progressCallback): Promise<void>;
}