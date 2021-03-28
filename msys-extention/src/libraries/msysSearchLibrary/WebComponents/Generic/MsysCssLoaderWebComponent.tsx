import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { SPComponentLoader } from '@microsoft/sp-loader';

/**
 * Guardare metodo getCustomWebComponents di MsysSearchLibrary
 * <msys-css-loader data-css-url=""></msys-css-loader>
 */
 export class MsysCssLoaderWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {        

        let props = this.resolveAttributes();

        if(props.cssUrl) {
            SPComponentLoader.loadCss(props.cssUrl + "?v=" + Math.floor(Math.random() * 10));
        }        

        const customComponent = <></>;
        ReactDOM.render(customComponent, this);
    }
}