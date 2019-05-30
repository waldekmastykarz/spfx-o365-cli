import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
// necessary to avoid TS errors in xterm.d.ts
const xterm = require('xterm');

import styles from './Office365CliWebPart.module.scss';
import * as strings from 'Office365CliWebPartStrings';
import { AadHttpClient, AadHttpClientConfiguration, HttpClientResponse } from '@microsoft/sp-http';
require('xterm/dist/xterm.css');

export interface IOffice365CliWebPartProps {
  container: string;
  containerGroup: string;
  resourceGroup: string;
  subscription: string;
}

export default class Office365CliWebPart extends BaseClientSideWebPart<IOffice365CliWebPartProps> {
  private azMgmtHttpClient: AadHttpClient;
  private term: any;
  private socket: WebSocket;

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (err: any) => void): void => {
      this.context.aadHttpClientFactory
        .getClient('https://management.azure.com/')
        .then((client: AadHttpClient): void => {
          this.azMgmtHttpClient = client;
          resolve();
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  public render(): void {
    this.term = new (xterm as any).Terminal({ convertEol: true });

    this.domElement.innerHTML = '<div></div>';
    this.term.open(this.domElement.getElementsByTagName('div')[0]);
    if (!this.properties.subscription) {
      this.term.writeln(`! Specify Subscription ID in the web part's configuration`);
      return;
    }
    if (!this.properties.resourceGroup) {
      this.term.writeln(`! Specify Resource Group Name in the web part's configuration`);
      return;
    }
    if (!this.properties.containerGroup) {
      this.term.writeln(`! Specify Container Group Name in the web part's configuration`);
      return;
    }
    if (!this.properties.container) {
      this.term.writeln(`! Specify Container Name in the web part's configuration`);
      return;
    }

    this.term.clear();
    this.term.writeln('Connecting to the CLI...');
    this.azMgmtHttpClient
      .post(`https://management.azure.com/subscriptions/${this.properties.subscription}/resourceGroups/${this.properties.resourceGroup}/providers/Microsoft.ContainerInstance/containerGroups/${this.properties.containerGroup}/containers/${this.properties.container}/exec?api-version=2018-10-01`, AadHttpClient.configurations.v1, {
        headers: {
          'content-type': 'application/json'
        },
        body: JSON.stringify({
          "command": "/bin/bash",
          "terminalSize": {
            "rows": this.term.rows,
            "cols": this.term.cols
          }
        })
      })
      .then((res: HttpClientResponse): Promise<{ error?: { message: string; }; password: string; webSocketUri: string; }> => {
        return res.json();
      })
      .then((res: { error?: { message: string; }; password: string; webSocketUri: string; }): void => {
        if (res.error) {
          this.term.writeln(`\x1B[31mERROR: ${res.error.message}\x1B[0m`);
          return;
        }

        this.socket = new WebSocket(res.webSocketUri);
        this.socket.onopen = (e) => {
          this.socket.send(res.password);
        };
        this.term.on('data', (data) => {
          this.socket.send(data);
        });
        this.socket.onmessage = (e) => {
          this.term.write(e.data);
        };
        this.term.focus();
      }, (err: any): void => {
        this.term.writeln(`\x1B[31mERROR: ${err.toString()}\x1B[0m`);
      });
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
                PropertyPaneTextField('subscription', {
                  label: 'Subscription ID'
                }),
                PropertyPaneTextField('resourceGroup', {
                  label: 'Resource Group Name'
                }),
                PropertyPaneTextField('containerGroup', {
                  label: 'Container Group Name'
                }),
                PropertyPaneTextField('container', {
                  label: 'Container Name'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges() {
    return true;
  }
}
