import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'FlowPomptCmdSetCommandSetCommandSetStrings';
import SPHttpClientResponse, { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from "@microsoft/sp-page-context";;
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFlowPomptCmdSetCommandSetCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
}

const LOG_SOURCE: string = 'FlowPomptCmdSetCommandSetCommandSet';

export default class FlowPomptCmdSetCommandSetCommandSet extends BaseListViewCommandSet<IFlowPomptCmdSetCommandSetCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized FlowPomptCmdSetCommandSetCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const command: Command | undefined = this.tryGetCommand("COMMAND_1")
    if(command){
      console.log("SPPermission: " + SPPermission.editListItems);
      command.visible = event.selectedRows.length >= 1 && this.context.pageContext.list.permissions.hasPermission(SPPermission.editListItems);
      console.log("HAI OH!")
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    var folderUrl: string;
    var fileUrl: string;
    var sourceUrl: string;
    var folder: string;
    var fullFile: string;
    var fileName: string;
    var docExt: string;
    const siteUrl: string = this.context.pageContext.web.absoluteUrl;
    const siteCollection: string = siteUrl.substring(siteUrl.lastIndexOf(".com") + 4, siteUrl.length);
    if (event.selectedRows.length > 0) {
      fileUrl = event.selectedRows[0].getValueByName("FileRef");
      sourceUrl = fileUrl.replace(siteCollection, "");
      fullFile = event.selectedRows[0].getValueByName("FileLeafRef");
      fileName = event.selectedRows[0].getValueByName("FileName");
      docExt = event.selectedRows[0].getValueByName("File_x0020_Type");
      docExt = `.${docExt}`;
      folder = sourceUrl.replace(`/${fullFile}`, "");
    }

    switch (event.itemId) {
      case 'COMMAND_1':
        if (docExt == ".url") {
          Dialog.alert(`Attention, URL Links cannot be moved to the OneJhpiego Library.`);
        }
        else {
          Dialog.alert(`Attention, ${fullFile} will be moved to the OneJhpiego Library.  You will be notified via email when complete, you can then refresh your page to see changes.`);
          this.flowPostRequest(siteUrl, sourceUrl, folder, fileName, docExt);
        }
        break;
      default:
        throw new Error('Unknown command');
    }
  }
  private flowPostRequest(siteUrl, sourceUrl, folder, fileName, docExt): void {
    ////Test Flow:
    //const postURL = "https://prod-27.southeastasia.logic.azure.com:443/workflows/6d78f2849aac4376ab6ab920bd8ef2f0/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=WZGmbY90Vlm-JD97SkDg12Rn-F-nXwrHvsDcfQyka_Y";
    ////Prod Flow:
    const postURL = "https://prod-53.westus.logic.azure.com:443/workflows/2b4f6e5add044c1f9f081f7431640cbe/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=J1jGYczuIhIG6LWRUXDfXBwKgnQcU2L0mZthfiMEsdw";
    const submiter: string = this.context.pageContext.user.loginName;

    //const destSite: string = "https://m365x244049.sharepoint.com/sites/DRCStaff";
    const destSite: string = "https://jhpiego.sharepoint.com/sites/onejhpiego-library";
    const destFlolder: string = "/Shared Documents";

    const body: string = JSON.stringify({
      "siteUrl": `${siteUrl}`,
      "sourceUrl": `${sourceUrl}`,
      "sourceFolder": `${folder}`,
      "destSite": `${destSite}`,
      "destFolder": `${destFlolder}`,
      "docName": `${fileName}`,
      "docExt": `${docExt}`,
      "submiter": `${submiter}`
    });

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');

    const httpClientOptions: IHttpClientOptions = {
      body: body,
      headers: requestHeaders
    };

    this.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      httpClientOptions)
      .then((response: HttpClientResponse) => {
        // Access properties of the response object. 
        console.log(`Status code: ${response.status}`);
        console.log(`Status text: ${response.statusText}`);
        //response.json() returns a promise so you get access to the json in the resolve callback.
        response.json().then((responseJSON: JSON) => {
          console.log(responseJSON);
        });
      });

  }
}
