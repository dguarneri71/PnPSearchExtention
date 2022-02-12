import * as React from 'react';
import styles from './TitleBorder.module.scss';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { renderToStaticMarkup } from "react-dom/server";
import { Border } from '.';
import { ITitleBorderProps } from "./ITitleBorderProps";
import * as DOMPurify from 'dompurify';
import { DomPurifyHelper } from '../../Helpers/DomPurifyHelper ';

let borderClass = mergeStyles({
    marginTop: 10
});

let bodyClass = mergeStyles({
    padding: 10
});

export class TitleBorder extends React.Component<ITitleBorderProps, {}> {
    private _domPurify: any;

    public constructor(props: ITitleBorderProps) {
        super(props);

        this._domPurify = DOMPurify;

        this._domPurify.setConfig({
            WHOLE_DOCUMENT: true
        });

        this._domPurify.addHook('uponSanitizeElement', DomPurifyHelper.allowCustomComponentsHook); //consente di inserire altri web component nel template
        this._domPurify.addHook('uponSanitizeAttribute', DomPurifyHelper.allowCustomAttributesHook); //consente di gestire attributi cutom nei web component

        console.log("Border - _domPurify: ", this._domPurify);
    }

    //Esempio di rendering con elementi vuoti <> usati al posto di React.Fragment
    public render(): React.ReactElement<ITitleBorderProps> {
        let textColor: string = this.props.textColor ? this.props.textColor : "#000";

        //Se non ho il bordo la classe del bordo l'ha do al contenitore interno
        let container_class: string = this.props.hide ? borderClass : "";

        console.log("Props:", this.props);
        console.log("Title Background Color:", this.props.titleBkgColor);

        let divStyle = {
            color: textColor,
            background: this.props.titleBkgColor ? this.props.titleBkgColor : this.props.color,
        } as React.CSSProperties;

        let icon = <span></span>;

        if (this.props.icon) {
            icon = <Icon iconName={this.props.icon} className={styles.componentIcon} />;
        }

        if (this.props.className) {
            borderClass = this.props.className;
        }

        if (this.props.bodyClassName) {
            bodyClass = this.props.bodyClassName;
        }

        let extraClasses = "";

        if (this.props.titleClassName) {
            extraClasses = this.props.titleClassName;
        }

        let content: JSX.Element = <div className={`${styles.componentContainer} ${container_class}`}>
            {!this.props.hideTitle &&
                <div className={`${styles.componentTitleBorder} ${extraClasses}`} style={divStyle}>
                    {icon}<span className={styles.componentText}>{this.props.title}</span>
                </div>
            }
            <div className={bodyClass} dangerouslySetInnerHTML={{ __html: this._domPurify.sanitize(this.props.contentTemplate) }}></div>
        </div>;

        let contentText = renderToStaticMarkup(content);

        return (
            <Border size={this.props.size} color={this.props.color} hide={this.props.hide} className={borderClass} contentTemplate={contentText} cssUrl={this.props.cssUrl} />
        );
    }
}