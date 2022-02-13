import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from './DownloadWebComponent.module.scss';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { IconButton, IIconProps, initializeIcons } from 'office-ui-fabric-react';
import { HttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IDataService } from '../../Classes/Services/IDataService';
import SPDataService from '../../Classes/Services/SPDataService';
import QueryData from '../../Classes/Entities/QueryData';
import DownloadFile from '../../Classes/Entities/DownloadFile';

export interface IDownloadComponentProps {
    content?: {}; //tutto il contenuto della Search Result WP
    label?: string;
    icon?: string;
    context: PageContext;
    httpClient: HttpClient;
}

export interface IDownloadComponentState {
    isCalloutVisible: boolean;
    callOutMsg: string;
}

// Initialize icons in case this example uses them
initializeIcons();

const LOG_SOURCE: string = 'DownloadComponent';
const downloadIcon: IIconProps = { iconName: 'Download' };

export class DownloadComponent extends React.Component<IDownloadComponentProps, IDownloadComponentState> {
    private dataService: IDataService;

    constructor(props: IDownloadComponentProps) {
        super(props);
        console.log(LOG_SOURCE + " - props: ", this.props);
        this.dataService = new SPDataService(this.props.context, this.props.httpClient);

        this.state = {
            isCalloutVisible: false,
            callOutMsg: ""
        };
    }

    public render() {
        console.log(LOG_SOURCE + " - Content: ", this.props.content);
        let items: any[] = this.props.content["data"]["items"];
        console.log(LOG_SOURCE + " - Items: ", items);
        let label: string = this.props.label;
        if (this.props.icon) {
            downloadIcon.iconName = this.props.icon;
        }
        return <>
            {
                this.props.label ?
                    (<span>{label} <IconButton iconProps={downloadIcon} title="Download All" ariaLabel="Download All" onClick={this.__download.bind(this)} /></span>)
                    :
                    (<IconButton iconProps={downloadIcon} title="Download All" ariaLabel="Download All" onClick={this.__download.bind(this)} />)
            }
        </>;
    }

    private __download(event): void {
        let downloadFiles: DownloadFile[] = [];
        let queryText: string = "*";
        if (this.props.content["inputQueryText"]) {
            queryText = this.props.content["inputQueryText"];
        }
        let count: number = this.props.content["data"]["totalItemsCount"];
        let webUrl: string = this.props.content["context"]["web"]["absoluteUrl"];
        let dataSourceProperties = this.props.content["properties"]["dataSourceProperties"];
        let enableQueryRules: boolean = dataSourceProperties["enableQueryRules"];
        let queryTemplate: string = dataSourceProperties["queryTemplate"];
        let resultSourceId: string = dataSourceProperties["resultSourceId"];
        let query: QueryData = new QueryData(queryText, enableQueryRules, queryTemplate, resultSourceId);
        console.log(LOG_SOURCE + " - Query: ", query);

        this.dataService.getSearchResult(query, count).then(results => {
            console.log(LOG_SOURCE + " - Search results: ", results);
            for (let index = 0; index < results.length; index++) {
                const element = results[index];
                if (element["FileType"] && element["FileExtension"] !== "aspx") {
                    let downloadFile: DownloadFile = new DownloadFile(element["Path"], element["Filename"], webUrl);
                    downloadFiles.push(downloadFile);
                }
            }
            console.log(LOG_SOURCE + " - DownloadFiles: ", downloadFiles);
            if (downloadFiles.length > 0) {
                this.download_files(downloadFiles);
            }
        });
    }

    private download_files(files: DownloadFile[]): void {
        function download_next(i: number) {
            if (i >= files.length) {
                return;
            }
            var a = document.createElement('a');
            a.href = files[i].webUrl + "/_layouts/download.aspx?SourceURL=" + encodeURI(files[i].download);
            console.log(LOG_SOURCE + " - download_next() - count: " + i + " - url", a.href);
            a.target = '_blank';
            a.setAttribute("data-interception", "off");
            // Use a.download if available, it prevents plugins from opening.
            if ('download' in a) {
                a.download = files[i].filename;
            }
            // Add a to the doc for click to work.
            (document.body || document.documentElement).appendChild(a);
            if (a.click) {
                a.click(); // The click method is supported by most browsers.
            }
            // Delete the temporary link.
            a.parentNode.removeChild(a);
            // Download the next file with a small timeout. The timeout is necessary
            // for IE, which will otherwise only download the first file.
            setTimeout(() => {
                download_next(i + 1);
            }, 500);
        }
        // Initiate the first download.
        download_next(0);
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
            const customComponent = <DownloadComponent context={_pageContext} httpClient={_httpClient} {...props} />;
            ReactDOM.render(customComponent, this);
        });
    }
}