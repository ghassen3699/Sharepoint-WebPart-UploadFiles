import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';
// import { escape } from '@microsoft/sp-lodash-subset';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import * as bootstrap from 'bootstrap';
require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');

import styles from './UploadFilesWebPart.module.scss';
import * as strings from 'UploadFilesWebPartStrings';

export interface IUploadFilesWebPartProps {
  description: string;
}

export default class UploadFilesWebPart extends BaseClientSideWebPart<IUploadFilesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div>
      <h3>Upload File</h3>
      <div>
        <input type="file" id="uploadFile" multiple="false"value="Select file" >
        <input type="button" id="uploadButton" value="Upload">
      </div>
    </div>
    `;
    this.uploadFileEventListener()
  }

  // upload file event listener
  private uploadFileEventListener(){
    const _absoluteURL = this.context.pageContext.web.absoluteUrl
    document.getElementById('uploadButton').addEventListener('click',() => this.UploadFile(_absoluteURL))
  }


  // upload file to a list
  private UploadFile(_absoluteURL){
    var fileUploaded = (<HTMLInputElement>document.getElementById('uploadFile')).files;
    for(let i = 0; i < fileUploaded.length; i++){
      var file = fileUploaded[i]

      var spOpts: ISPHttpClientOptions = {
        headers: {
          "Accept": "application/json",
          "Content-type": "application/json"
        },
        body:file
      }

      var url= `${_absoluteURL}/_api/Web/Lists/getByTitle('Documents')/RootFolder/Files/Add(url='${file.name}', overwrite=true)`;
      this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,spOpts)
      .then((response:SPHttpClientResponse)=>{
        response.json().then((responseJson)=>{
          console.log(responseJson.Name)
        })
      })
    }
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
