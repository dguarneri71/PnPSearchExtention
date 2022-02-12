import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { TitleBorder } from '../../Components/Graphics';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ITitleBorderWebComponentProps } from './ITitleBorderWebComponentProps';

export class TitleBorderWebComponent extends React.Component<ITitleBorderWebComponentProps, null> {
    public render() {
        console.log("TitleBorderWebComponent - SIZE: ", this.props.size);
        console.log("TitleBorderWebComponent - COLOR: ", this.props.color);
        console.log("TitleBorderWebComponent - HIDE: ", this.props.hide);
        console.log("TitleBorderWebComponent - CLASS-NAME: ", this.props.className);
        console.log("TitleBorderWebComponent - CONTENT-TEMPLATE: ", this.props.contentTemplate);

        return <TitleBorder
            size={this.props.size}
            color={this.props.color}
            hide={this.props.hide}
            hideTitle={this.props.hideTitle}
            className={this.props.className}
            icon={this.props.icon}
            textColor={this.props.textColor}
            title={this.props.title}
            titleClassName={this.props.titleClassName}
            bodyClassName={this.props.bodyClassName} 
            titleBkgColor={this.props.titleBkgColor}
            contentTemplate={this.props.contentTemplate} />;
    }
}

/**
 * Guardare metodo getCustomWebComponents di MsysSearchLibrary
 * <msys-title-border data-color="" data-title-bkg-color="" data-text-color="" data-size="" data-hide="" data-hide-title="" data-icon="" data-title=""  data-class-name="" data-title-class-name="" data-body-class-name="" data-css-url="">
 *  <template id="border-content">
 *       {content}
 *   </template>
 * </msys-title-border>
 */
export class MsysTitleBorderWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        console.log("MsysTitleBorderWebComponent props:", props);

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
        
        const customComponent = <TitleBorderWebComponent {...props} contentTemplate={contentTemplateContent} />;
        ReactDOM.render(customComponent, this);
    }
}