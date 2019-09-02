import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import styles from './SpfxAadWebPart.module.scss';
import * as strings from 'SpfxAadWebPartStrings';
import { HttpClient } from '@microsoft/sp-http';

export interface ISpfxAadWebPartProps {
  ClientID: string;
  APIUrl: string;
}

export default class SpfxAadWebPart extends BaseClientSideWebPart<ISpfxAadWebPartProps> {
  private error: string = null;
  private result: string = null;
  private token: string = null;
  private decodedToken: any = null;

  public async onInit(): Promise<void> {
    if (!this.properties.ClientID || this.properties.ClientID === "00000000-0000-0000-0000-000000000000") {
      this.error = "Please set Client ID property in the webpart settings";
      return;
    }

    try {
      // Get token PROVIDER
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      // Get TOKEN
      this.token = await tokenProvider.getToken(this.properties.ClientID, false);
    }
    catch(error) {
      this.error = error;
    }

    if (this.token) {
      try {
        // Call secured API with token
        let headers = new Headers();
        headers.set('Authorization', `Bearer ${this.token}`);
        const response = await this.context.httpClient.get(this.properties.APIUrl, HttpClient.configurations.v1, { headers });
        if (response.ok) {
          const result = await response.json();
          this.result = result.message;
          this.decodedToken = result.decodedToken;
        }
        else {
          throw new Error(`${response.status}: ${response.statusText}`);
        }
      }
      catch(error) {
        this.error = `Error calling API URL: ${this.properties.APIUrl}. ${error}`;
      }
    }
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spfxAad }">
        <div class="${ styles.container }">
          <h2>Result</h2>
          ${this.error
            ? `<div class="${ styles.error }">${this.error}</div>`
            : `<div class="${ styles.result }">${this.result}</div>`
          }
          ${this.decodedToken
            ?  `<h2>Decoded JWT</h2>
                <table "${ styles.decodedToken }">
                  <tr><td><b>Claim</b></td><td><b>Value</b></td></tr>
                  ${Object.keys(this.decodedToken).map(k => `<tr><td>${k}</td><td>${this.decodedToken[k]}</td></tr>`).join("")}
               </table>`
            : ''
          }
          ${this.token
            ? `<h2>Raw JWT</h2>
               <div class="${ styles.token }">${this.token}</div>`
            : ''
          }
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.GroupName,
              groupFields: [
                PropertyPaneTextField('ClientID', {
                  label: strings.ClientIDFieldLabel
                }),
                PropertyPaneTextField('APIUrl', {
                  label: strings.APIUrlFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
