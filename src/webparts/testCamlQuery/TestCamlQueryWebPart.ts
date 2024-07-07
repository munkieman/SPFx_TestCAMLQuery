import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestCamlQueryWebPart.module.scss';
import * as strings from 'TestCamlQueryWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; 
//import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { LogLevel, PnPLogging } from "@pnp/logging";
require("bootstrap");

let libCount : number = 0;

export interface ITestCamlQueryWebPartProps {
  description: string;
  division : string;
  teamTermID : string;
  parentTermID : string;
  libraryName: string[];
  libraries : string[];
  libraryNamePrev : string;
  dataResults: any[];
}

export default class TestCamlQueryWebPart extends BaseClientSideWebPart<ITestCamlQueryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public async render(): Promise<void> {

    this.properties.libraryNamePrev = "";
    this.domElement.innerHTML = `
    <section class="${styles.testCamlQuery} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div class="${styles.row}">     
        <div class="${styles.row}" id="libraryName"></div>
      </div>
      <div class="row">
        <div class="text-black" id="testFolders"></div>
        <div class="col-auto" id="policiesFolders"></div>
        <div class="col ms-2 filesContainer" id="policiesFiles"></div>
      </div>               
    </section>`;

    //*** check data exists for libraries and display tab buttons */
    await this._getDataAsync(false,"","").then( () =>{
      setTimeout(async () => {
        await this._renderlibraryTabsAsync("");        
      }, 1000);      
    });

    //await this._renderListAsync(); //.then( () => {});
      //this._libraryListeners()
  }

  private async _getDataAsync(flag:boolean,library:string,category:string): Promise<void> {
    console.log('getDataAsync',flag,library,category);
    const dcDivisions : string[] = ["asm","cen","cnn","emp","hea"];

    try{

      if(library===""){
        this.properties.libraryName = ["Policies", "Procedures","Guides", "Forms", "General"];
      }else{
        this.properties.libraryName = [library];
      }

      // *** clear dataResults array
      this.properties.dataResults = [];
      
      // *** loop through all libraries in the libraryName array
      for (let x = 0; x < this.properties.libraryName.length; x++) {

        // *** check through all division DCs for data
        dcDivisions.forEach(async (site,index)=>{
          this._getData(flag,site,this.properties.libraryName[x],"IPES Wales",category)
          .then(async (response) => {
            console.log("getDataAsync",response,flag);
            if(response.length>0){
              if(!flag){
                await this._setLibraryTabs(this.properties.libraryName[x]);        
              }else{                   
                await this.addToResults(response).then(async ()=>{            
                  //await this._renderFolders(this.properties.libraryName[x]).then(async () => {
                    //await this.setFolderListeners(this.properties.libraryName[x]);              
                  //});
                }); 
              }
            }
          })
          .catch((err) => {console.log("error:",err)});
        });
      }
    } catch (err) {
      //await this.addError(this.properties.teamName,"getLibraryTabs",err);
      //Log.error('DocumentCentre', new Error('getLibraryTabs Error message'), err);
    }
    return;
  }

  private async _getData(flag:boolean,site:string,library:string,team:string,category:string): Promise<any> {      
    console.log('checkdata');

    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));  
    const tenant_uri = this.context.pageContext.web.absoluteUrl.split('/',3)[2];
    const dcTitle = site+"_dc";
    const webDC = Web([sp.web,`https://${tenant_uri}/sites/${dcTitle}/`]); 
    let rowLimitString : string;
    let view: string = "";

    if(!flag){
      rowLimitString="<RowLimit>10</RowLimit>";
    }else{
      rowLimitString="";
    }
    
    if (category === "") {
      view =
        `<View>
        <Query>
          <Where>
            <Or>
              <Eq>
                <FieldRef Name="DC_Team"/>
                <Value Type="TaxonomyFieldType">${team}</Value>
              </Eq>
              <Contains>
                <FieldRef Name="DC_SharedWith"/>
                <Value Type="TaxonomyFieldTypeMulti">${team}</Value>
              </Contains>
            </Or>
          </Where>
          <OrderBy>
            <FieldRef Name="DC_Division" Ascending="TRUE" />
            <FieldRef Name="DC_Folder" Ascending="TRUE" />
            <FieldRef Name="DC_SubFolder01" Ascending="TRUE" />
            <FieldRef Name="DC_SubFolder02" Ascending="TRUE" />
            <FieldRef Name="DC_SubFolder03" Ascending="TRUE" />
            <FieldRef Name="LinkFilename" Ascending="TRUE" />
          </OrderBy>          
        </Query>
        ${rowLimitString}
      </View>`;
    } else {
      view =
        `<View>
          <Query>
            <Where>
              <Or>
                <Eq>
                  <FieldRef Name="DC_Category"/>
                  <Value Type="TaxonomyFieldType">${category}</Value>
                </Eq>
                <Contains>
                  <FieldRef Name="DC_SharedWith"/>
                  <Value Type="TaxonomyFieldTypeMulti">${category}</Value>
                </Contains>
              </Or>
            </Where>
            <OrderBy>
              <FieldRef Name="DC_Division" Ascending="TRUE" />
              <FieldRef Name="DC_Folder" Ascending="TRUE" />
              <FieldRef Name="DC_SubFolder01" Ascending="TRUE" />
              <FieldRef Name="DC_SubFolder02" Ascending="TRUE" />
              <FieldRef Name="DC_SubFolder03" Ascending="TRUE" />
              <FieldRef Name="LinkFilename" Ascending="TRUE" />
            </OrderBy>           
          </Query>
          ${rowLimitString}
        </View>`;
    }

    return webDC.lists.getByTitle(library)
      .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
      .then(async (response) => {
        //console.log("checkdata",response);
        return response;  
      })
      .catch(() => {});    
  }

  private async addToResults(results:any):Promise<void>{
    let count:number=0; 
    console.log("results",this.properties.dataResults);

    if(this.properties.dataResults !== undefined){count=this.properties.dataResults.length}

    //if(this.properties.teamTerm!==undefined){
      results.forEach(async (item: any, index: number) => {
    //    console.log("teamTermID",this.properties.teamTermID,"team TermGuid",item.DC_Team.TermGuid);

    //    if(item.DC_Team.TermGuid===this.properties.teamTermID){
          this.properties.dataResults[count]=item;
    //      count++;
    //    }
      });
    //}
    console.log("addToResults results",this.properties.dataResults,"count",count);
    return;
  }

  private async _renderlibraryTabsAsync(category:string): Promise<void> {
    //if(this.properties.libraries !== undefined){
      this.properties.libraries.sort()
    //}
      for(let x=0; x<this.properties.libraries.length;x++){     
        console.log("libraryTabsAsync",this.properties.libraries[x],x);
        await this._renderLibraryTabs(this.properties.libraries[x]).then( async ()=> {
          this._setLibraryListeners();
          // *** get custom tabs from termstore and add library column
          //await this.renderCustomTabsAsync();              
        });  
      }
    
    return;
  }

  private async _setLibraryTabs(library: string): Promise<void>{
    console.log("setLibrary",library,this.properties.libraryNamePrev,libCount);
    
    if(this.properties.libraryNamePrev !== library){
      this.properties.libraryNamePrev = library;
      this.properties.libraries[libCount] = library;
      libCount++;
    }

    return;
  }

  private async _renderLibraryTabs(library:string): Promise<void> {
    console.log('renderlist');

    //const dataTarget:string=library.toLowerCase();
    let html: string = '';

    //console.log("renderlist",library);
    html = `<button class="btn btn-primary text-center mb-1" id="${library}_btn" type="button"><h6 class="libraryText">${library}</h6></button>`;

    if(this.domElement.querySelector('#libraryName') !== null){
      this.domElement.querySelector('#libraryName')!.innerHTML += html;
    }
    return;
  }
  
  private _setLibraryListeners() : void {
    console.log("librarylisteners");

    try {
      let timer:any;
      let libraryBtn : string = "";

      //console.log("libraries",this.properties.libraries);

      for(let lib=0;lib<this.properties.libraries.length;lib++){
        libraryBtn = `${this.properties.libraries[lib]}_btn`;
        
        const elem=document.getElementById(libraryBtn);
        elem?.addEventListener("click", event => {
          if(event.detail===1){
            timer=setTimeout(async () => {
              console.log("setLibraryListeners click",this.properties.libraries[lib]);
              await this._getDataAsync(true,this.properties.libraries[lib],""); 
            },400);
          }        
        });

        elem?.addEventListener("dblclick",event => {
          clearTimeout(timer);
        });        
      }      
    } catch (err) {
      //await this.addError(this.properties.teamName, "setLibraryListeners", err);
      //Log.error('DocumentCentre', new Error('setLibraryListeners Error message'), err);
    }
    return;    
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    this.properties.libraries = [];
    this.properties.libraryName = [];

    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");

    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
