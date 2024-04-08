import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebComponent, IDataFilter, IDataFilterConfiguration } from '@pnp/modern-search-extensibility';
import { IconButton, IIconProps, initializeIcons, CommandBarButton, Spinner, DefaultButton, PrimaryButton, Dialog, DialogFooter, DialogType, ProgressIndicator } from 'office-ui-fabric-react';
import { HttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IDataService } from '../../Classes/Services/IDataService';
import SPDataService from '../../Classes/Services/SPDataService';
import QueryData from '../../Classes/Entities/QueryData';
import { FieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-controls-react/lib/FieldCollectionData';
import { Helper } from '../../Helpers/Helper';
import { FieldCollectionDataValue } from '../../Classes/Entities/FieldCollectionDataValue';
import { stringIsNullOrEmpty } from "@pnp/common";
import { DateHelper } from '../../Helpers/DateHelper';

export interface IExcelComponentProps {
    content?: {}; //tutto il contenuto della Search Result WP
    label?: string;
    icon?: string;
    labelsListTitle?: string;
    context: PageContext;
    httpClient: HttpClient;
    dateHelper: DateHelper;
}

export interface IExcelComponentState {
    isCalloutVisible: boolean;
    callOutMsg: string;
    showPanel: boolean;
    headerLabels: any[];
    errorMsg: string;
    showSpinner: boolean;
    percentComplete: number;
    totalResultsCount: number;
    partialResultsCount: number;
}

// Initialize icons in case this example uses them
initializeIcons();

const LOG_SOURCE: string = 'ExcelComponent';
const icon: IIconProps = { iconName: 'ExcelDocument' };
const LABEL: string = 'Download Results in Excel';
const HEARDER_TITLE: string = "Title";
const HEARDER_DISPLAY_NAME: string = "DisplayName";
const HEARDER_TYPE: string = "Type";
const HEARDER_FORMAT: string = "Format";
const STORAGE_KEY: string = "ExcelHeaders";

//TODO: aggiungere salvataggio in cookie delle headerLabels
export class ExcelComponent extends React.Component<IExcelComponentProps, IExcelComponentState> {
    private dataService: IDataService;
    private moment: any;

    constructor(props: IExcelComponentProps) {
        super(props);
        console.log(LOG_SOURCE + " - props: ", this.props);
        this.dataService = new SPDataService(this.props.context, this.props.httpClient);
        let _values: any[] = [];

        if (_values.length == 0) {
            let dataSourceProperties = this.props.content["properties"]["dataSourceProperties"];
            let properties: string[] = dataSourceProperties["selectedProperties"];
            _values = properties.map((value, index, array) => {
                let valueData: FieldCollectionDataValue = new FieldCollectionDataValue();
                valueData.Title = value;
                valueData.DisplayName = value;
                return valueData;
            });
        }

        console.log(LOG_SOURCE + " - HeaderLabels: ", _values);

        this.state = {
            isCalloutVisible: false,
            callOutMsg: "",
            showPanel: false,
            headerLabels: _values,
            errorMsg: null,
            showSpinner: false,
            percentComplete: 0,
            totalResultsCount: 0,
            partialResultsCount: 0
        };
    }

    public async componentWillMount() {
        this.moment = await this.props.dateHelper.moment();

        if (stringIsNullOrEmpty(this.props.labelsListTitle) == false) {
            try {
                const items: FieldCollectionDataValue[] = await this.dataService.getLabels(this.props.labelsListTitle);
                console.log(LOG_SOURCE + " - componentWillMount() - HeaderLabels: ", items);
                if (items.length > 0) {
                    this.setState({
                        headerLabels: items
                    });
                }
            }
            catch (e) {
                console.log(LOG_SOURCE + " - componentWillMount() - Error: ", e);
                let msg = "List '" + this.props.labelsListTitle + "' does not exist.";

                // https://pnp.github.io/pnpjs/concepts/error-handling/
                // are we dealing with an HttpRequestError?
                if (e?.isHttpRequestError) {
                    // we can read the json from the response
                    const json = await e.response.json();
                    console.log(LOG_SOURCE + " - componentWillMount() - JSON Error: ", json);
                    // if we have a value property we can show it
                    msg = typeof json["odata.error"] === "object" ? json["odata.error"].message.value : e.message;
                }

                this.setState({
                    errorMsg: msg
                });
            }
        }
    }

    public render() {
        console.log(LOG_SOURCE + " - render() - state: ", this.state);
        const { percentComplete, partialResultsCount, totalResultsCount } = this.state;
        let label: string = this.props.label ? this.props.label : LABEL;
        if (this.props.icon) {
            icon.iconName = this.props.icon;
        }

        let countMsg = partialResultsCount + " / " + totalResultsCount;

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
                {(stringIsNullOrEmpty(this.props.labelsListTitle) || !stringIsNullOrEmpty(this.state.errorMsg)) && <FieldCollectionData
                    key={"FieldCollectionData"}
                    label={"Define Excel Columns"}
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
                }

                {this.state.showSpinner && <ProgressIndicator label="Waiting..." description={countMsg} percentComplete={percentComplete} />}

                {!stringIsNullOrEmpty(this.state.errorMsg) && <div style={{ color: "red", fontSize: "12px", paddingTop: "10px" }}>{this.state.errorMsg}</div>}

                <DialogFooter>
                    <PrimaryButton disabled={this.state.showSpinner} onClick={this.__download.bind(this)} text="Download Excel" />
                    <DefaultButton onClick={this._hidePanel} text="Cancel" />
                </DialogFooter>
            </Dialog>
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
        this.setState({
            headerLabels: values
        });
    }

    private __download(event): void {
        this.setState({
            showSpinner: true
        });

        let selectedProperties = Helper.getValuesForArray(this.state.headerLabels, HEARDER_TITLE);
        console.log(LOG_SOURCE + " - selectedProperties: ", selectedProperties);
        let headers = Helper.getValuesForArray(this.state.headerLabels, HEARDER_DISPLAY_NAME);
        console.log(LOG_SOURCE + " - headers: ", headers);
        let types = Helper.getValuesForArray(this.state.headerLabels, HEARDER_TYPE);
        console.log(LOG_SOURCE + " - types: ", types);
        let formats = Helper.getValuesForArray(this.state.headerLabels, HEARDER_FORMAT);
        console.log(LOG_SOURCE + " - formats: ", formats);

        let query: QueryData = new QueryData(this.props, selectedProperties);

        console.log(LOG_SOURCE + " - Query: ", query);

        this.dataService.getSearchResult(query, this.moment, this.progressCallback.bind(this)).then(results => {
            console.log(LOG_SOURCE + " - Search results: ", results);
            Helper.downloadExcel(results, selectedProperties, headers, types, formats);

            this.setState({
                showPanel: false,
                showSpinner: false
            });
        }).catch(reason => {
            this.setState({
                showSpinner: false,
                errorMsg: JSON.stringify(reason)
            });
        });
    }

    private progressCallback(percentComplete: number, partial: number, total: number) {
        console.log(LOG_SOURCE + " - progressCallback() - Percentuale: ", percentComplete);
        console.log(LOG_SOURCE + " - progressCallback() - Totale parziale: ", partial);
        console.log(LOG_SOURCE + " - progressCallback() - Totale: ", total);
        this.setState({
            percentComplete: percentComplete,
            partialResultsCount: partial,
            totalResultsCount: total
        });
    }
}

//<msys-results-excel data-label="Download Excel" data-content="{{JSONstringify this 2}}" data-icon="ExcelLogo" labels-list-title="Header Mapping"></msys-results-excel>
export class ExcelWebComponent extends BaseWebComponent {
    public constructor() {
        super();
    }

    public async connectedCallback() {
        let props = this.resolveAttributes();
        this._serviceScope.whenFinished(() => {
            console.log(LOG_SOURCE + " - _serviceScope: ", this._serviceScope);
            let _httpClient: HttpClient = this._serviceScope.consume(HttpClient.serviceKey);
            console.log(LOG_SOURCE + " - _httpClient: ", _httpClient);
            let _pageContext: PageContext = this._serviceScope.consume(PageContext.serviceKey);
            console.log(LOG_SOURCE + " - _pageContext: ", _pageContext);
            let _dateHelper = this._serviceScope.consume<DateHelper>(DateHelper.ServiceKey);
            const customComponent = <ExcelComponent context={_pageContext} httpClient={_httpClient} dateHelper={_dateHelper} {...props} />;
            ReactDOM.render(customComponent, this);
        });
    }
}