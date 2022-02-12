import * as React from 'react';
import { IBorderProps } from "./IBorderProps";
import * as DOMPurify from 'dompurify';
import { DomPurifyHelper } from '../../Helpers/DomPurifyHelper ';

export class Border extends React.Component<IBorderProps, {}> {
    private _domPurify: any;

    public constructor(props: IBorderProps) {
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
    //Definizione e utilizzo della proprietà className
    public render(): React.ReactElement<IBorderProps> {


        if (!document.getElementById) document.write('<link rel="stylesheet" type="text/css" href="' + this.props.cssUrl + '">');

        return (<>
            {
                !this.props.hide ?
                    (
                        <div
                            style={{ borderStyle: "solid", borderWidth: this.props.size + "px", borderColor: this.props.color }}
                            className={this.props.className} dangerouslySetInnerHTML={{ __html: this._domPurify.sanitize(this.props.contentTemplate) }}>
                            {this.props.children}
                        </div>
                    ) :
                    (
                        <div className={this.props.className} dangerouslySetInnerHTML={{ __html: this._domPurify.sanitize(this.props.contentTemplate) }}></div>
                    )
            }
        </>);
    }
}