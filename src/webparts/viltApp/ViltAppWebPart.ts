import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ViltAppWebPart.module.scss';
import * as strings from 'ViltAppWebPartStrings';

export interface IViltAppWebPartProps {
  description: string;
}

export default class ViltAppWebPart extends BaseClientSideWebPart <IViltAppWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.viltApp }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a id="openDialog" href="#" class="${ styles.button }">
                <span class="${ styles.label }">Open Dialog</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    `;
    this._setDialogButton(); 
  }

  private _setDialogButton(): void {
    // get the button
    var btn = document.getElementById("openDialog")
    // if in Microsoft Teams
    if (this.context.sdks.microsoftTeams) {
      let taskInfo = {
        title: null,
        height: null,
        width: null,
        url: null,
        card: null,
        fallbackUrl: null,
        completionBotId: null,
      };
    
      taskInfo.url = "https://aristolabs.sharepoint.com/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3Fteams%26personal%26componentId=b60ee32c-1d57-452a-9c22-e124e83c6d71%26forceLocale=en-us";
      taskInfo.title = "Custom Dialog";
      taskInfo.height = "large";
      taskInfo.width = "large";
      
      // Btn handler
      btn.onclick = () => {
        console.log("Attempting to open the dialog");
        this.context.sdks.microsoftTeams.teamsJs.tasks.startTask(taskInfo, (err, result) => { console.log("Submit handler"); });
      }
    } else {
      btn.onclick = () => {
        alert("Not running in Teams");
      }
    }




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
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
              })
            ]
          }
        ]
      }
    ]
  };
}
}
