import * as React from 'react';
import { IBorderProps } from "./IBorderProps";
import * as DOMPurify from 'dompurify';

export class Border extends React.Component<IBorderProps, {}> {    
    //Esempio di rendering con elementi vuoti <> usati al posto di React.Fragment
    //Definizione e utilizzo della proprietà className
    public render(): React.ReactElement<IBorderProps> {
        

        if (!document.getElementById) document.write('<link rel="stylesheet" type="text/css" href="'+ this.props.cssUrl +'">');

        return (<>
            {
                !this.props.hide ?
                    (
                        <div
                            style={{ borderStyle: "solid", borderWidth: this.props.size + "px", borderColor: this.props.color }}
                            className={this.props.className} dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(this.props.contentTemplate) }}>
                            {this.props.children}
                        </div>
                    ) :
                    (
                        <div className={this.props.className} dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(this.props.contentTemplate) }}></div>
                    )
            }
        </>);
    }
}