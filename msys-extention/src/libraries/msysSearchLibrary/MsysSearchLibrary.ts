import { IExtensibilityLibrary, IComponentDefinition, ISuggestionProviderDefinition } from "@pnp/modern-search-extensibility";
import { MyCustomComponentWebComponent } from "./WebComponents/Demo/CustomComponent";
import { DGDemoComponentWebComponent } from "./WebComponents/Demo/DemoWebComponent";
import { MsysBorderWebComponent,  MsysTitleBorderWebComponent, MsysCssLoaderWebComponent } from "./WebComponents/Generic";
import * as Handlebars from 'handlebars';
import { isEmpty } from "@microsoft/sp-lodash-subset";
import { DownloadWebComponent, ExcelWebComponent, InvokeFlowWebComponent } from "./WebComponents/Interactive";

export class MsysSearchLibrary implements IExtensibilityLibrary {
  //Registro i custom Web Components
  //TODO fare componente per scaricare risultati come Excel o CSV
  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'msys-custom-component',
        componentClass: MyCustomComponentWebComponent
      },
      {
        componentName: 'msys-demo-component',
        componentClass: DGDemoComponentWebComponent
      },
      {
        componentName: 'msys-border',
        componentClass: MsysBorderWebComponent
      },
      {
        componentName: 'msys-title-border',
        componentClass: MsysTitleBorderWebComponent
      },
      {
        componentName: 'msys-css-loader',
        componentClass: MsysCssLoaderWebComponent
      },
      {
        componentName: 'msys-download-all',
        componentClass: DownloadWebComponent
      },
      {
        componentName: 'msys-call-flow',
        componentClass: InvokeFlowWebComponent
      },
      {
        componentName: 'msys-results-excel',
        componentClass: ExcelWebComponent
      }
    ];
  }

  //Registro i custom Suggestion Providers
  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  //Registro i custom Heandelbars Helpers
  public registerHandlebarsCustomizations(namespace: typeof Handlebars) {    
    // Return format string
    // Usage: {{cleanText 'value'}}
    namespace.registerHelper("cleanText", (text: string) => {
      if (!isEmpty(text)) {
        return text.replace(" ", "-").replace("&", "").toLowerCase().replace("à", "a").replace("è", "e").replace("é", "e").replace("ì", "i").replace("ù", "u");
      }
    });

    // Return the URL of the search result item
    // Usage: <a href="{{getUrl item}}">
    namespace.registerHelper("getUrl", (item: any) => {
      if (item && item.ServerRedirectedURL && item.Path) {
        return !isEmpty(item.ServerRedirectedURL) ? item.ServerRedirectedURL : item.Path;
      }
    });

    // Return .....
    // Usage: <a href="{{getLocateUrl item}}">
    namespace.registerHelper("getLocateUrl", (item: any) => {
      if (item && item.ParentLink && item.SPWebUrl && item.ListItemID) {
        var parentLink = item.ParentLink;
        var relParentLink = parentLink.replace(item.SPWebUrl + "/", "");
        var relParentLinkArray = relParentLink.split("/");
        var url = item.SPWebUrl + "/" + relParentLinkArray[0] + "/Forms/All%20documents.aspx?FilterType1=Text&FilterField1=ID&FilterValue1=" + item.ListItemID;
        return url;
      }
      else {
        return "";
      }
    });

    // Return .....
    // Usage: <a href="{{getDispUrl item}}">
    namespace.registerHelper("getDispUrl", (item: any) => {
      if (item && item.ParentLink && item.SPWebUrl && item.ListItemID) {
        var parentLink = item.ParentLink;
        var relParentLink = parentLink.replace(item.SPWebUrl + "/", "");
        var relParentLinkArray = relParentLink.split("/");
        var url = item.SPWebUrl + "/" + relParentLinkArray[0] + "/Forms/DispForm.aspx?ID=" + item.ListItemID;
        return url;
      }
      else {
        return "";
      }
    });
  }
}