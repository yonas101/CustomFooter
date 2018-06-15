import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";
import { MSGraphClient } from "@microsoft/sp-client-preview";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import * as strings from "CustomFooterApplicationCustomizerStrings";
import styles from "./CustomFooterStyling.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as $ from "jquery";

export interface ICustomFooterApplicationCustomizerProperties {
  Bottom: string;
  currentMailId: string;
}

export default class CustomFooterApplicationCustomizer extends BaseApplicationCustomizer<ICustomFooterApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent;
  //Endpoint for the call against MS Graph, use ms graph explorer to change to get your specific data.
  public graphEndpoint: string = "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta";
  //public graphEndpoint: string = "https://graph.microsoft.com/v1.0/me/messages";
  public currentMailId: string = "";

  @override
  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceHolders
    );
    this._renderPlaceHolders();
    this._graphCall();
    return Promise.resolve<void>();
  }

  private _renderPlaceHolders(): void {

    //Handling the placeholder
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );

      //Return if placeholder is not valid/available
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }

      if (this.properties) {
        if (this._bottomPlaceholder.domElement) {
          this._bottomPlaceholder.domElement.innerHTML = `
        <div id="bottomFooterExpanded" class="${styles.bottomFooterExpanded} ms-bgColor-themeDark ms-fontColor-white">
        <div id="outlookIcon" class="${styles.outlookIcon}"><i class="ms-Icon ms-Icon--OutlookLogo" aria-hidden="true"></i></div>
        <div id="footerText" class="${styles.footerText}"></div>
        <button id="isReadButton" class="${styles.isReadButton} ms-bgColor-themeDark ms-fontColor-white" >Markera som läst</button>
        <button id="deleteButton" class="${styles.deleteButton} ms-bgColor-themeDark ms-fontColor-white" >Ta bort</button>
        <div class="${styles.footerRightIcons}"><i id="closeButton" class="ms-Icon ms-Icon--ChevronDownEnd6" aria-hidden="true"></i></div>
        </div>
        <style>
        .od-UserFeedback-button{display: none;}
        </style>`;
        }
      }
    }

    //Custom actions too make unbind on the dual clickevent triggered.
    $("#closeButton").unbind("click");
    //Action to style the footer according to if its minimized or not.
    $("#closeButton").click(() => {
      $("#closeButton").toggleClass("ms-Icon--ChevronUpEnd6");
      var footerContainer = $("#bottomFooterExpanded");
      var footerTextContainer = $("#footerText");
      var footerOutlookContainer = $("#outlookIcon");
      var footerReadBtn = $("#isReadButton");
      var footerDelBtn = $("#deleteButton");
      footerContainer.toggleClass(styles.bottomFooterMinimized);
      footerTextContainer.toggleClass(styles.footerTextMinimized);
      footerOutlookContainer.toggleClass(styles.outlookIconMinimized);
      footerReadBtn.toggleClass(styles.isReadButtonMinimized);
      footerDelBtn.toggleClass(styles.deleteButtonMinimized);
    });

      //CUSTOM CODE FOR isReadButton
      $("#isReadButton").unbind("click");
      $("#isReadButton").click(() => {
       this._graphSetMailIsRead();
      });

      //CUSTOM CODE FOR deleteButton
      $("#deleteButton").unbind("click");
      $("#deleteButton").click(() => {
        this._graphDeleteMail();
      });

  }

  private _onDispose(): void {
    console.log("Disposed custom bottom placeholder.");
  }

  //Function that makes the rest call against MS Graph and handles the response.
  private async _graphCall() {
    const aadClient: AadHttpClient = new AadHttpClient(this.context.serviceScope, "https://graph.microsoft.com");
    let response: HttpClientResponse = await aadClient.get(this.graphEndpoint, AadHttpClient.configurations.v1);

    if (response.ok) {
      response.json().then((result) => {
        var i;
        for (i = 0; i < result.value.length; i++) {
          console.log("Result:" + result.value[i].subject);
          if (result.value[i].isRead === false) {
            //get id of currentmail that could later be updated or deleted

            this.currentMailId = result.value[i].id;
            console.log("current mail id:" + result.value[i].id);
            document.getElementById("footerText").innerHTML =
              "Sender: " +
              result.value[i].sender.emailAddress.address +
              "</br>" +
              "Subject: " +
              result.value[i].subject +
              "</br>" +
              "Body: " +
              result.value[i].bodyPreview;
            break;
          } else {
            document.getElementById("footerText").innerHTML =
              "Inga olästa mail just nu.";
          }
        }
      });
    } else {
      let error = new Error("error:" + response.statusText);
      throw error;
    }
  }

  //DEVELOPEMENT OF FEATURE "SET MAIL AS READ"
  private _graphSetMailIsRead() {
    const client: MSGraphClient = this.context.serviceScope.consume(MSGraphClient.serviceKey);

    const mailIsRead : MicrosoftGraph.Message = {
      isRead: true
    };

     client
      .api('me/messages/'+this.currentMailId)
      .patch(mailIsRead)
      .then((mailResponse) => {
        console.log(mailResponse);
   }).then(()=>{
    this._graphCall();
   });
  }

  private _graphDeleteMail() {
    const client: MSGraphClient = this.context.serviceScope.consume(MSGraphClient.serviceKey);

    console.log("Id of mail to delete:" + this.currentMailId);

     client
      .api('me/messages/'+ this.currentMailId)
      .delete()
      .then((mailResponse) => {
        console.log(mailResponse);
   }).then(()=>{
     this._graphCall();
   });
  }
}
