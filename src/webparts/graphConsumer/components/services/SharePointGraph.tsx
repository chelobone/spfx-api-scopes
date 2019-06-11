import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { MSGraphClientFactory, MSGraphClient } from '@microsoft/sp-http';
import { AadHttpClientFactory, AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export interface ICustomGraphService {
    RootWeb(): Promise<JSON>;
    executeMyRequest(): Promise<JSON>;
}

export class CustomGraphService implements ICustomGraphService {

    //Create a ServiceKey which will be used to consume the service.
    public static readonly serviceKey: ServiceKey<ICustomGraphService> =
        ServiceKey.create<ICustomGraphService>('spfx-api-scopes:ICustomGraphService', CustomGraphService);

    private _msGraphClientFactory: MSGraphClientFactory;
    private _aadHttpClientFactory: AadHttpClientFactory;

    constructor(serviceScope: ServiceScope) {
        //console.log(serviceScope)
        serviceScope.whenFinished(() => {
            this._msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
            this._aadHttpClientFactory = serviceScope.consume(AadHttpClientFactory.serviceKey);

        });
    }

    public RootWeb(): Promise<JSON> {
        return new Promise<JSON>((resolve, reject) => {
            this._msGraphClientFactory.getClient().then((_msGraphClient: MSGraphClient) => {
                //alert(_msGraphClient.api('/me'))
                _msGraphClient.api('/sites/root').get((error, site: any, rawResponse?: any) => {
                   
                    if (error) {
                        alert("GraphService" + error.message)
                    }

                    resolve(site);
                });
            });
        });
    }

    public executeMyRequest(): Promise<JSON> {
        //You can add your own AAD resource here. Using the Graph API resource for simplicity.
        return this._aadHttpClientFactory.getClient("https://graph.microsoft.com").then((_aadHttpClient: AadHttpClient) => {

            //This would be your custom endpoint
            return _aadHttpClient.get('https://graph.microsoft.com/v1.0/me', AadHttpClient.configurations.v1).then((response: HttpClientResponse) => {
                return response.json();
            }).catch((e)=>{
                alert(e);
                return JSON.parse(null);
            });
        });
    }
}