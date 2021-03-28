import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { Border } from '../../Components/Graphics';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IBorderWebComponentProps {
    color?: string;
    size?: number;
    hide?: boolean;
    className?: string;
    //cssUrl?: string;   

    contentTemplate: string;
}

export class BorderWebComponent extends React.Component<IBorderWebComponentProps, null> {
    public render() {
        return <Border size={this.props.size} color={this.props.color} hide={this.props.hide} className={this.props.className} contentTemplate={this.props.contentTemplate} />;
    }
}

/**
 * Guardare metodo getCustomWebComponents di MsysSearchLibrary
 * <msys-border data-color="" data-size="" data-hide="" data-class-name="" data-css-url=""></msys-border>
 */
export class MsysBorderWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {        

        let props = this.resolveAttributes();

        if(props.cssUrl) {
            SPComponentLoader.loadCss(props.cssUrl + "?v=" + Math.floor(Math.random() * 10));
        }

        const domParser = new DOMParser();
        const htmlContent: Document = domParser.parseFromString(this.innerHTML, 'text/html');

        // Get the templates
        const contentTemplate = htmlContent.getElementById('border-content');

        let contentTemplateContent = null;

        if (contentTemplate) {
            contentTemplateContent = contentTemplate.innerHTML;
        }

        const customComponent = <BorderWebComponent {...props} contentTemplate={contentTemplateContent} />;
        ReactDOM.render(customComponent, this);
    }
}