import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphConsumerWebPartStrings';
import GraphConsumer from './components/GraphConsumer';
import { IGraphConsumerProps } from './components/IGraphConsumerProps';
import { ClientMode } from './components/ClientMode';
import * as microsoftTeams from '@microsoft/teams-js';
// import {
//   CustomGraphService
// } from '../graphConsumer/components/services/SharePointGraph';
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { MSGraphClient, GraphHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export interface IGraphConsumerWebPartProps {
  clientMode: ClientMode;
}

export default class GraphConsumerWebPart extends BaseClientSideWebPart<IGraphConsumerWebPartProps> {
  private _teamsContext: microsoftTeams.Context;


  public render(): void {
    const element: React.ReactElement<IGraphConsumerProps> = React.createElement(
      GraphConsumer,
      {
        clientMode: this.properties.clientMode,
        context: this.context,
      }
    );

//     this.context.graphHttpClient
//     .get("v1.0/sites/root", GraphHttpClient.configurations.v1)
//     .then((res:HttpClientResponse)=>{
//       //console.log(res.json());
// var response = res.json();
//       //alert(res);
//       response.then((r:any)=>{
//         console.log(r.displayName);
//         alert(r.displayName);
//       }).catch((e)=>{
//         alert(e);
//       });
//       //alert("OK")
//     }).catch((err)=>{
//       alert(err);
//     });
    // .getClient()
    // .then((client: MSGraphClient): void => {
    //   //console.log(client);
    //   //alert(client);
    //   client
    //     //.api("/sites/imagen.sharepoint.com,5467ed13-ab7f-40fb-b5e4-e570adac0785,f7852e05-c564-4801-856a-5c88760f5ade/lists/cb6ef9de-cf38-4507-80e6-92e9409b7858/items?expand=Descargable")
    //     .api("/groups/{"+this._teamsContext.teamId+"}/members")
    //     .version("v1.0")
    //     //.select("id,webUrl,createdBy,lastModifiedBy,createdDateTime,Descargable")
    //     //.expand("Descargable")
    //     //.filter(`(givenName eq '${escape(this.state.searchFor)}') or (surname eq '${escape(this.state.searchFor)}') or (displayName eq '${escape(this.state.searchFor)}')`)
    //     .get((err, res:any) => {
    //       if (err) {
    //         //console.error(err);
    //         alert("WebPart err " + err.message)
    //         //this.setState({ errorDialog: false, loadComplete: true, errorMessage: err.message + " " + err.statusCode })
    //         return;
    //       }
    //       console.log(res);
    //       alert("WebPart " + res);
    //     });
    // });
    // const _customGraphServiceInstance = this.context.serviceScope.consume(CustomGraphService.serviceKey);
    
    // _customGraphServiceInstance.RootWeb().then((user: any) => {
    //  //alert(user);
    // }); 

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('clientMode', {
                  label: strings.ClientModeLabel,
                  options: [
                    { key: ClientMode.aad, text: "AadHttpClient" },
                    { key: ClientMode.graph, text: "MSGraphClient" },
                  ]
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
