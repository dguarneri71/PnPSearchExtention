import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { IconButton, IIconProps, initializeIcons, CommandBarButton, Panel, PanelType, DefaultButton, PrimaryButton, Dialog, DialogFooter, DialogType } from 'office-ui-fabric-react';
import { HttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IDataService } from '../../Classes/Services/IDataService';
import SPDataService from '../../Classes/Services/SPDataService';
import QueryData from '../../Classes/Entities/QueryData';
import { FieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-controls-react/lib/FieldCollectionData';
import { Helper } from '../../Helpers/Helper';
import { FieldCollectionDataValue } from '../../Classes/Entities/FieldCollectionDataValue';
import { PnPClientStorage } from "@pnp/core";
import { stringIsNullOrEmpty } from "@pnp/common";

export interface IExcelComponentProps {
    content?: {}; //tutto il contenuto della Search Result WP
    label?: string;
    icon?: string;
    context: PageContext;
    httpClient: HttpClient;
}

export interface IExcelComponentState {
    isCalloutVisible: boolean;
    callOutMsg: string;
    showPanel: boolean;
    headerLabels: any[];
}

// Initialize icons in case this example uses them
initializeIcons();

const LOG_SOURCE: string = 'ExcelComponent';
const icon: IIconProps = { iconName: 'ExcelDocument' };
const LABEL: string = 'Download Results in Excel';
const HEARDER_TITLE: string = "Title";
const HEARDER_DISPLAY_NAME: string = "DisplayName";
const STORAGE_KEY: string = "ExcelHeaders";

//TODO: aggiungere salvataggio in cookie delle headerLabels
export class ExcelComponent extends React.Component<IExcelComponentProps, IExcelComponentState> {
    private dataService: IDataService;
    private storage: PnPClientStorage;

    constructor(props: IExcelComponentProps) {
        super(props);
        console.log(LOG_SOURCE + " - props: ", this.props);
        this.dataService = new SPDataService(this.props.context, this.props.httpClient);
        let _values: any[] = [];
        this.storage = new PnPClientStorage();
        _values = this.storage.local.get(STORAGE_KEY);
        console.log(LOG_SOURCE + " - Storage HeaderLabels: ", _values);

        if (_values == null) {
            let dataSourceProperties = this.props.content["properties"]["dataSourceProperties"];
            let properties: string[] = dataSourceProperties["selectedProperties"];
            _values = properties.map((value, index, array) => {
                let valueData: FieldCollectionDataValue = new FieldCollectionDataValue();
                valueData.Title = value;
                valueData.DisplayName = value;
                return valueData;
            });
            this.storage.local.put(STORAGE_KEY, _values);
        }

        console.log(LOG_SOURCE + " - HeaderLabels: ", _values);

        this.state = {
            isCalloutVisible: false,
            callOutMsg: "",
            showPanel: false,
            headerLabels: _values
        };
    }

    public render() {
        let label: string = this.props.label ? this.props.label : LABEL;
        if (this.props.icon) {
            icon.iconName = this.props.icon;
        }

        return <>
            {
                this.props.label ?
                    (<CommandBarButton iconProps={icon} text={label} ariaLabel={label} onClick={this._showPanel.bind(this)} />)
                    :
                    (<IconButton iconProps={icon} title={label} ariaLabel={label} onClick={this._showPanel.bind(this)} />)
            }

            <Dialog
                hidden={!this.state.showPanel}
                onDismiss={this._hidePanel}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Download Search Result',
                    subText: ''
                  }}
                  modalProps={{
                    isBlocking: true,
                    styles: { main: { maxWidth: 450 } }
                  }}
            >
                <FieldCollectionData
                    key={"FieldCollectionData"}
                    label={"Excel Columns"}
                    manageBtnLabel={"Manage"}
                    onChanged={this._onChangeFieldCollectionData.bind(this)}
                    panelHeader={"Define Excel File Columns"}
                    enableSorting={true}
                    executeFiltering={(searchFilter: string, item: any) => {
                        return item[HEARDER_TITLE] === +searchFilter;
                    }}
                    itemsPerPage={10}
                    fields={[
                        { id: HEARDER_TITLE.valueOf(), title: "Property Name", type: CustomCollectionFieldType.string, required: true },
                        { id: HEARDER_DISPLAY_NAME.valueOf(), title: "Column Name", type: CustomCollectionFieldType.string, required: true },
                    ]}
                    value={this.state.headerLabels}
                />
                <DialogFooter>
                    <PrimaryButton onClick={this.__download.bind(this)} text="Download Excel" />
                    <DefaultButton onClick={this._hidePanel} text="Cancel" />
                </DialogFooter>
            </Dialog>

            {/* <Panel
                isOpen={this.state.showPanel}
                onDismiss={this._hidePanel}
                type={PanelType.largeFixed}
                closeButtonAriaLabel="Close"
                headerText="Download Search Result"
            >
                <DefaultButton text="Download Excel" onClick={this.__download.bind(this)} />
                <FieldCollectionData
                    key={"FieldCollectionData"}
                    label={"Excel Columns"}
                    manageBtnLabel={"Manage"}
                    onChanged={this._onChangeFieldCollectionData.bind(this)}
                    panelHeader={"Define Excel File Columns"}

                    executeFiltering={(searchFilter: string, item: any) => {
                        return item[HEARDER_TITLE] === +searchFilter;
                    }}
                    itemsPerPage={10}
                    fields={[
                        { id: HEARDER_TITLE.valueOf(), title: "Property Name", type: CustomCollectionFieldType.string, required: true },
                        { id: HEARDER_DISPLAY_NAME.valueOf(), title: "Column Name", type: CustomCollectionFieldType.string, required: true },
                    ]}
                    value={this.state.headerLabels}
                />
            </Panel> */}
        </>;
    }

    private _showPanel = () => {
        this.setState({ showPanel: true });
    }

    private _hidePanel = () => {
        this.setState({ showPanel: false });
    }

    private _onChangeFieldCollectionData(values: any[]): void {
        console.log(LOG_SOURCE + " - FieldCollectionData: ", values);
        this.storage.local.put(STORAGE_KEY, values);
        this.setState({
            headerLabels: values
        });
    }

    private __download(event): void {
        let queryText: string = "*";
        if (this.props.content["inputQueryText"]) {
            queryText = this.props.content["inputQueryText"];
        }
        let count: number = this.props.content["data"]["totalItemsCount"];
        let dataSourceProperties = this.props.content["properties"]["dataSourceProperties"];
        let enableQueryRules: boolean = dataSourceProperties["enableQueryRules"];
        let queryTemplate: string = dataSourceProperties["queryTemplate"];
        let resultSourceId: string = dataSourceProperties["resultSourceId"];

        let selectedProperties = Helper.getValuesForArray(this.state.headerLabels, HEARDER_TITLE);
        console.log(LOG_SOURCE + " - selectedProperties: ", selectedProperties);
        let headers = Helper.getValuesForArray(this.state.headerLabels, HEARDER_DISPLAY_NAME);
        console.log(LOG_SOURCE + " - headers: ", headers);

        let query: QueryData = new QueryData(queryText, enableQueryRules, queryTemplate, resultSourceId, selectedProperties);

        console.log(LOG_SOURCE + " - Query: ", query);

        this.dataService.getSearchResult(query, count).then(results => {
            console.log(LOG_SOURCE + " - Search results: ", results);
            Helper.downloadExcel(results, selectedProperties, headers);
        });
    }
}

//<msys-results-excel data-label="Download Excel" data-content="{{JSONstringify this 2}}" data-icon="ExcelLogo"></msys-results-excel>
export class ExcelWebComponent extends BaseWebComponent {
    public constructor() {
        super();
    }

    public async connectedCallback() {
        let props = this.resolveAttributes();
        this._serviceScope.whenFinished(() => {
            let _httpClient: HttpClient = this._serviceScope.consume(HttpClient.serviceKey);
            console.log(LOG_SOURCE + " - _httpClient: ", _httpClient);
            let _pageContext: PageContext = this._serviceScope.consume(PageContext.serviceKey);
            console.log(LOG_SOURCE + " - _pageContext: ", _pageContext);
            const customComponent = <ExcelComponent context={_pageContext} httpClient={_httpClient} {...props} />;
            ReactDOM.render(customComponent, this);
        });
    }
}