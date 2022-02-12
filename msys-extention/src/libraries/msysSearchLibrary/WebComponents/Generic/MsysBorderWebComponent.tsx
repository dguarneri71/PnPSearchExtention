import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { Border } from '../../Components/Graphics';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { IBorderWebComponentProps } from './IBorderWebComponentProps';

export class BorderWebComponent extends React.Component<IBorderWebComponentProps, null> {    
    public render() {
        console.log("BorderWebComponent - SIZE: ", this.props.size);
        console.log("BorderWebComponent - COLOR: ", this.props.color);
        console.log("BorderWebComponent - HIDE: ", this.props.hide);
        console.log("BorderWebComponent - CLASS-NAME: ", this.props.className);
        console.log("BorderWebComponent - CONTENT-TEMPLATE: ", this.props.contentTemplate);

        return <Border size={this.props.size} color={this.props.color} hide={this.props.hide} className={this.props.className} contentTemplate={this.props.contentTemplate} />;
    }
}

/**
 * Guardare metodo getCustomWebComponents di MsysSearchLibrary
 * <msys-border data-color="" data-size="" data-hide="" data-class-name="" data-css-url="">
 *   <template id="border-content">
 *       {content}
 *   </template>
 * </msys-border>
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