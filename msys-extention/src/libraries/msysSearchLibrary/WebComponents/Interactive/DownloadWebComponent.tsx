import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from './DownloadWebComponent.module.scss';
import { BaseWebComponent, IDataFilter, IDataFilterConfiguration } from '@pnp/modern-search-extensibility';
import { IconButton, IIconProps, initializeIcons, CommandBarButton, Spinner, Dialog, DialogType, ProgressIndicator } from 'office-ui-fabric-react';
import { HttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IDataService } from '../../Classes/Services/IDataService';
import SPDataService from '../../Classes/Services/SPDataService';
import QueryData from '../../Classes/Entities/QueryData';
import DownloadFile from '../../Classes/Entities/DownloadFile';
import { DateHelper } from '../../Helpers/DateHelper';
import { stringIsNullOrEmpty } from "@pnp/common";

export interface IDownloadComponentProps {
    content?: {}; //tutto il contenuto della Search Result WP
    label?: string;
    icon?: string;
    context: PageContext;
    httpClient: HttpClient;
    dateHelper: DateHelper;
}

export interface IDownloadComponentState {
    isCalloutVisible: boolean;
    callOutMsg: string;
    showPanel: boolean;
    showSpinner: boolean;
    errorMsg: string;
    spinnerMsg: string;
    percentComplete: number;
    filename: string;
}

// Initialize icons in case this example uses them
initializeIcons();

const LOG_SOURCE: string = 'DownloadComponent';
const downloadIcon: IIconProps = { iconName: 'Download' };
const LABEL: string = 'Download All';
const INTERVAL_INCREMENT = 0.01;

export class DownloadComponent extends React.Component<IDownloadComponentProps, IDownloadComponentState> {
    private dataService: IDataService;
    private moment: any;

    constructor(props: IDownloadComponentProps) {
        super(props);
        console.log(LOG_SOURCE + " - props: ", this.props);
        this.dataService = new SPDataService(this.props.context, this.props.httpClient);

        this.state = {
            isCalloutVisible: false,
            callOutMsg: "",
            showPanel: false,
            showSpinner: false,
            errorMsg: null,
            spinnerMsg: "Waiting...",
            percentComplete: 0,
            filename: ""
        };
    }

    public async componentWillMount() {
        this.moment = await this.props.dateHelper.moment();
    }

    public render() {
        const { percentComplete, filename } = this.state;
        console.log(LOG_SOURCE + " - Content: ", this.props.content);
        let items: any[] = this.props.content["data"]["items"];
        console.log(LOG_SOURCE + " - Items: ", items);
        let label: string = this.props.label ? this.props.label : LABEL;
        if (this.props.icon) {
            downloadIcon.iconName = this.props.icon;
        }
        return <>
            {
                this.props.label ?
                    (<CommandBarButton iconProps={downloadIcon} text={label} ariaLabel={label} onClick={this.__download.bind(this)} />)
                    :
                    (<IconButton iconProps={downloadIcon} title={label} ariaLabel={label} onClick={this.__download.bind(this)} />)
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
                {this.state.showSpinner && <Spinner label={this.state.spinnerMsg} />}

                <ProgressIndicator label="Zipping file" description={filename} percentComplete={percentComplete} />

                {!stringIsNullOrEmpty(this.state.errorMsg) && <div style={{ color: "red", fontSize: "12px", paddingTop: "10px" }}>{this.state.errorMsg}</div>}
            </Dialog>
        </>;
    }

    private _showPanel = () => {
        this.setState({ showPanel: true });
    }

    private _hidePanel = () => {
        this.setState({ showPanel: false });
    }

    private __download(event): void {
        let downloadFiles: DownloadFile[] = [];
        let webUrl: string = this.props.content["context"]["web"]["absoluteUrl"];
        let query: QueryData = new QueryData(this.props, ["Filename", "FileType", "FileExtension", "Path", "SPSiteUrl"]);

        console.log(LOG_SOURCE + " - Query: ", query);
        
        this.dataService.getSearchResult(query, this.moment, this.voidSearchCallback.bind(this)).then(results => {
            console.log(LOG_SOURCE + " - Search results: ", results);
            for (let index = 0; index < results.length; index++) {
                const element = results[index];
                if (element["FileType"] && element["FileExtension"] !== "aspx") {
                    let downloadFile: DownloadFile = new DownloadFile(element["Path"], element["Filename"], webUrl, element["SPSiteUrl"]);
                    downloadFiles.push(downloadFile);
                }
            }
            console.log(LOG_SOURCE + " - DownloadFiles: ", downloadFiles);
            if (downloadFiles.length > 0) {
                this.setState({
                    showPanel: true,
                    showSpinner: false
                });
                this.download_files_zip(downloadFiles);
            }
        });
    }

    private voidSearchCallback(percentComplete: number, partial: number, total: number) {
        console.log(LOG_SOURCE + " - progressCallback() - Percentuale: ", percentComplete);
        console.log(LOG_SOURCE + " - progressCallback() - Totale parziale: ", partial);
        console.log(LOG_SOURCE + " - progressCallback() - Totale: ", total);
    }

    private download_files_zip(files: DownloadFile[]): void {
        this.dataService.downloadZipFile(files, this.progressCallback.bind(this));
    }

    private progressCallback(files: DownloadFile[], index: number, msg: string, finish: boolean): void {
        if (finish) {
            this.setState({
                showPanel: false,
                showSpinner: false
            });
        } else {
            var perc = ((index + 1) / files.length) + INTERVAL_INCREMENT;
            this.setState({
                percentComplete: perc,
                filename: files[index].filename,
                errorMsg: msg
            });
        }
    }
}

//<msys-download-all data-label="Download All" data-content="{{JSONstringify this 2}}" data-icon="Download"></msys-download-all>
export class DownloadWebComponent extends BaseWebComponent {
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
            let _dateHelper = this._serviceScope.consume<DateHelper>(DateHelper.ServiceKey);
            const customComponent = <DownloadComponent context={_pageContext} httpClient={_httpClient} dateHelper={_dateHelper} {...props} />;
            ReactDOM.render(customComponent, this);
        });
    }
}