import * as React from 'react';
import styles from './DemoWebComponent.module.scss';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { DefaultButton, PrimaryButton, Stack, IStackTokens } from 'office-ui-fabric-react';
import { PanelBtn } from '../../Components/Containers/Panels/PanelBtn';

export interface IDemoWebComponentProps {
    /**
     * A sample object param: data-my-object-param
     */
    myObjectParam?: any;

    /**
     * A sample string param
     */
    cssUrl?: string;
}

export interface IDemoWebComponenState {
}

const stackTokens: IStackTokens = { childrenGap: 40 };

export class DemoWebComponent extends React.Component<IDemoWebComponentProps, IDemoWebComponenState> {
    public render() {        
        console.log("DG-Lib: myObjectParam", this.props.myObjectParam);
        const myObject: any = this.props.myObjectParam;        
        return <div>
            <Stack horizontal tokens={stackTokens}>
                <PanelBtn cssUrl={this.props.cssUrl} item={myObject}>{this.props.children}</PanelBtn>
            </Stack>
        </div>;
    }
}

/**
 * Guardare metodo getCustomWebComponents di MsysSearchLibrary
 * <msys-demo-component data-my-object-param="" data-css-url=""></msys-demo-component>
 */
export class DGDemoComponentWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        const customComponent = <DemoWebComponent {...props} />;
        ReactDOM.render(customComponent, this);
    }
}