import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from './InvokeFlowWebComponent.module.scss';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { Text, Callout, DirectionalHint, IconButton, IIconProps, initializeIcons, CommandBarButton } from 'office-ui-fabric-react';
import { HttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { stringIsNullOrEmpty } from "@pnp/common";
import { IDataService } from '../../Classes/Services/IDataService';
import SPDataService from '../../Classes/Services/SPDataService';
import { Constants } from '../../Classes/Constants';
import InvokeFlowResult from '../../Classes/Entities/InvokeFlowResult';

export interface IInvokeFlowComponentProps {
    content?: {}; //tutto il contenuto della Search Result WP
    label?: string;
    icon?: string;
    settingsKey?: string;
    listSettings?: string;
    context: PageContext;
    httpClient: HttpClient;
}

export interface IInvokeFlowComponentState {
    flowUrl: string;
    isCalloutVisible: boolean;
    callOutMsg: string;
    hasError: boolean;
}

// Initialize icons in case this example uses them
initializeIcons();

const LOG_SOURCE: string = 'InvokeFlowComponent';
const SUCCESS_INVOKE_FLOW: string = "Process running. You will receive an email at the end of the job.";
const ERROR_INVOKE_FLOW: string = "Action failed, please retry.";
const ERROR_CONFIGURATION: string = "Please configure web component.";
const invokeIcon: IIconProps = { iconName: 'Download' };
const errorIcon: IIconProps = { iconName: 'Warning' };
const LABEL: string = 'Download All';

export class InvokeFlowComponent extends React.Component<IInvokeFlowComponentProps, IInvokeFlowComponentState> {
    private dataService: IDataService;

    constructor(props: IInvokeFlowComponentProps) {
        super(props);
        console.log(LOG_SOURCE + " - props: ", this.props);
        this.dataService = new SPDataService(this.props.context, this.props.httpClient);

        let _hasError: boolean = false;
        if ("listSettings" in this.props && "settingsKey" in this.props) {
            if (stringIsNullOrEmpty(this.props.listSettings) || stringIsNullOrEmpty(this.props.settingsKey)) {
                _hasError = true;
            }
        }
        else {
            _hasError = true;
        }

        this.state = {
            flowUrl: null,
            isCalloutVisible: false,
            callOutMsg: "",
            hasError: _hasError
        };
    }

    public componentDidMount(): void {
        let configListName: string = this.props.listSettings;
        if (stringIsNullOrEmpty(configListName) === false && stringIsNullOrEmpty(this.state.flowUrl)) {     
            this.dataService.getSettingsBySpecificKey(this.props.listSettings, this.props.settingsKey).then(values => {
                console.log(LOG_SOURCE + " - componentDidMount() - settings: ", values);
                for (let index = 0; index < values.length; index++) {
                    const element = values[index];
                    console.log(LOG_SOURCE + " - componentDidMount() - item: ", element);
                    if (element.title == Constants.INVOKE_FLOW_URL_PARAM) {
                        console.log(LOG_SOURCE + " - componentDidMount() - setting value: ", element.value);
                        this.setState({
                            flowUrl: element.value
                        });
                        return;
                    }
                }
            });
        }
    }

    public render() {
        const { isCalloutVisible, hasError } = this.state;
        console.log(LOG_SOURCE + " - Content: ", this.props.content);
        let items: any[] = this.props.content["data"]["items"];
        console.log(LOG_SOURCE + " - Items: ", items);
        let label: string = this.props.label ? this.props.label : LABEL;
        if (this.props.icon) {
            invokeIcon.iconName = this.props.icon;
        }

        let classID = "InvokeFlow_Callout";

        return <>
            <span className={`${classID}`}>
                {!hasError &&
                    /* (this.props.label ?
                        (<span>{label} <IconButton iconProps={invokeIcon} title={label} ariaLabel={label} onClick={this.__invoke.bind(this)} /></span>)
                        :
                        (<IconButton iconProps={invokeIcon} title={label} ariaLabel={label} onClick={this.__invoke.bind(this)} />)
                    ) */
                    (
                        this.props.label ?
                            (<CommandBarButton iconProps={invokeIcon} text={label} ariaLabel={label} onClick={this.__invoke.bind(this)} />)
                            :
                            (<IconButton iconProps={invokeIcon} title={label} ariaLabel={label} onClick={this.__invoke.bind(this)} />)
                    )
                }
                {hasError &&
                    <IconButton iconProps={errorIcon} title={ERROR_CONFIGURATION} ariaLabel={ERROR_CONFIGURATION} onClick={this.__showError.bind(this)} />
                }
            </span>
            {isCalloutVisible && (
                <Callout
                    className={styles.callout}
                    role="alertdialog"
                    gapSpace={0}
                    isBeakVisible={true}
                    beakWidth={16}
                    target={`.${classID}`}
                    onDismiss={() => this.setState({ isCalloutVisible: false })}
                    directionalHint={DirectionalHint.rightCenter}
                    setInitialFocus
                    calloutMaxWidth={180}
                >
                    <div className={styles.inner} style={{ padding: "10px", minHeight: "30px" }}>
                        <Text>
                            {this.state.callOutMsg}
                        </Text>
                    </div>
                </Callout>
            )}
        </>;
    }

    private __showError(event): void {
        this.setState({
            isCalloutVisible: true,
            callOutMsg: ERROR_CONFIGURATION
        });
    }

    private async __invoke(event): Promise<void> {
        let queryText: string = "*";
        let userEmail: string = queryText = this.props.context.user.email;
        if (this.props.content["inputQueryText"]) {
            queryText = this.props.content["inputQueryText"];
        }
        const params: any = {
            "siteUrl": this.props.context.web.absoluteUrl,
            "data": queryText,
            "userEmail": userEmail
        };
        console.log(LOG_SOURCE + " - __invoke() - Flow URL: ", this.state.flowUrl);
        let result: InvokeFlowResult = await this.dataService.invokePowerAutomateFlowExtended(this.state.flowUrl, params, false);
        console.log(LOG_SOURCE + " - __invoke() - Flow Result: ", result);
        if (!result.success) {
            console.log(LOG_SOURCE + " - __invoke() - Flow Error: ", result.error);
            this.setState({
                isCalloutVisible: true,
                callOutMsg: ERROR_INVOKE_FLOW
            });
        } else {
            this.setState({
                isCalloutVisible: true,
                callOutMsg: SUCCESS_INVOKE_FLOW
            });
        }
    }
}

//<msys-call-flow data-content="{{JSONstringify this 2}}" data-label="Download All with Flow" data-list-settings="Settings" data-settings-key="SPFX_InvokeFlowComponent" data-icon="Page"></msys-call-flow>
export class InvokeFlowWebComponent extends BaseWebComponent {
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
            const customComponent = <InvokeFlowComponent context={_pageContext} httpClient={_httpClient} {...props} />;
            ReactDOM.render(customComponent, this);
        });
    }
}