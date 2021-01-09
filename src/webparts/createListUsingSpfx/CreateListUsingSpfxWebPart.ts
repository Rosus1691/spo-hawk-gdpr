import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CreateListUsingSpfxWebPartStrings';
import styles from './components/CreateListUsingSpfx.module.scss';
import CreateListUsingSpfx from './components/CreateListUsingSpfx';
import { ICreateListUsingSpfxProps } from './components/ICreateListUsingSpfxProps';
import { Web } from 'sp-pnp-js';
//commit test

export interface ICreateListUsingSpfxWebPartProps {
  description: string;
}

export default class CreateListUsingSpfxWebPart extends BaseClientSideWebPart<ICreateListUsingSpfxWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `

    <div class="${styles.createListUsingSpfx}">  
    
    <div class="${styles.container}">  
    
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
    
    <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
    
    <span class="ms-font-xl ms-fontColor-white" style="font-size:28px">Welcome to SPFx learning (create list using PnP JS library)</span>  
    
    <p class="ms-font-l ms-fontColor-white" style="text-align: left">Demo : Create SharePoint List in SPO using SPFx</p>  
    
    </div>  
    
    </div>  
    
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
    
    <div data-role="main" class="ui-content">    
    
                 <div>    
                  <input id="listTitle"  placeholder="List Name"/>  
                  <button id="createNewCustomListToSPO"  type="submit" >Create List</button>   
                 </div>    
    
               </div> 
    
    <br>  
    
    <div id="ListCreationStatusInSPOnlineUsingSPFx" />  
    
    </div>  
    
    </div>  
    
    </div>`;

    this.AddEventListeners();

  }

  private AddEventListeners(): void {
    document.getElementById('createNewCustomListToSPO').addEventListener('click', () => this.CreateListInSPOUsinPnPSPFx());
  }

  private CreateListInSPOUsinPnPSPFx(): void {
    let myWeb = new Web(this.context.pageContext.web.absoluteUrl);
    console.log("my web ::"+myWeb.toUrl.toString);
   

    //let mySPFxListTitle = "CustomList_using_SPFx_Framework"; 

    let mySPFxListTitle = document.getElementById('listTitle')["value"];

    let mySPFxListDescription = "Custom list created using the SPFx Framework";

    let listTemplateID = 100;

    let boolEnableCT = false;

    myWeb.lists.add(mySPFxListTitle, mySPFxListDescription, listTemplateID, boolEnableCT).then(function (splist) {

      document.getElementById("ListCreationStatusInSPOnlineUsingSPFx").innerHTML += `The SPO new list ` + mySPFxListTitle + ` has been created successfully using SPFx Framework.`;

    });

    myWeb.lists.getByTitle('My List').items.add
     
    //const r = list.select('Id');
    //console.log(r); 

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
