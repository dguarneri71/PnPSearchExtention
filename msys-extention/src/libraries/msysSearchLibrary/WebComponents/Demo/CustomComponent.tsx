import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';

export interface ICustomComponentProps {
    content?: {};
}

export interface ICustomComponenState {
}

export class CustomComponent extends React.Component<ICustomComponentProps, ICustomComponenState> {

    public render() {

        // Parse custom object
        console.log("Content: ", this.props.content);
        let items:any[] = this.props.content["data"]["items"];
        console.log("Items: ", items);
        let count = items.length;
        let currentPage = this.props.content["paging"]["currentPageNumber"];

        return <div>
            <span>Ciao: {count} - page: {currentPage}</span>
        </div>;
    }
}

//<msys-custom-component data-content="{{JSONstringify this 2}}"/>
export class MyCustomComponentWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        const customComponent = <CustomComponent {...props} />;
        ReactDOM.render(customComponent, this);
    }
}