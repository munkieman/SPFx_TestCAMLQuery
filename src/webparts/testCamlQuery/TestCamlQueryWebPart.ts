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
import { spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; 
//import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { LogLevel, PnPLogging } from "@pnp/logging";

//let libraryFlags: boolean[];

export interface ITestCamlQueryWebPartProps {
  description: string;
  division : string;
  teamTermID : string;
  parentTermID : string;
  libraryName: string[];
  flags: boolean[];
}

export default class TestCamlQueryWebPart extends BaseClientSideWebPart<ITestCamlQueryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public async render(): Promise<void> {
    const libraryName : string[] = ["Policies", "Procedures", "Guides", "Forms", "General"];
    let libraryNameHTML : string = "";
    let libraryFlagHTML : string = "";

    //libraryFlags=[];

    this.domElement.innerHTML = `
    <section class="${styles.testCamlQuery} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>Flags=${this.properties.flags}</div>
      <div class="${styles.row}">     
        <div class="${styles.column}" id="libraryName"><h3>Library Names</h3></div>
        <div class="${styles.column}" id="libraryFlag"><h3>Library Flags</h3></div>
      </div>
   
    </section>`;

    //console.log("libraryFlag array",libraryFlag);
    //console.log("libraryName array",this.properties.libraryName);
    //const libraryTabs : Element | null = this.domElement.querySelector("#librarynames");
    //console.log(this.properties.libraryFlag);

    //for(let x=0;x<this.properties.libraryFlag.length;x++){
      //console.log("libraryFlag",x);
      //libraryTabs!.innerHTML+=`<h4>${this.properties.libraryName[x]}</h4>`;
    //}

    for (let x = 0; x < libraryName.length; x++) {
      libraryNameHTML += `<div>${libraryName[x]}</div>`;
      libraryFlagHTML += `<div>${this.properties.flags[x]}</div>`;

      this.checkData(x,libraryName[x],"IPES Wales","");   
    }

    console.log("render",this.properties.flags); 

    if(this.domElement.querySelector('#libraryName') !== null){
      this.domElement.querySelector('#libraryName')!.innerHTML += libraryNameHTML;
    }  

    if(this.domElement.querySelector('#libraryFlag') !== null) {
      this.domElement.querySelector('#libraryFlag')!.innerHTML += libraryFlagHTML;
    }
    
  }

  private checkData(num:number,library:string,team:string,category:string): void { 
    
    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));  
    const tenant_uri = this.context.pageContext.web.absoluteUrl.split('/',3)[2];
    const dcDivisions : string[] = ["asm","cen","cnn","emp","hea"];
    let view: string = "";
    //let count:number=0; 

    if (category === "") {
      view =
        `<View><Query>
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
        </Query>
        </View>`;
    } else {
      view =
        `<View><Query>
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
        </Query>
        </View>`;
    }

    dcDivisions.forEach(async (site,index)=>{
      const dcTitle = site+"_dc";
      const webDC = Web([sp.web,`https://${tenant_uri}/sites/${dcTitle}/`]); 
      await webDC.lists.getByTitle(library)
        .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
        .then(async (results) => {
          console.log("check data",dcTitle," library",library,num,results);
          if(results.length>0){   
            console.log("response length",results.length); 
            this.setFlag(num,true);
          }else{
            this.setFlag(num,false);
          }             
        })
    });

    return;
    //console.log("library",this.properties.libraryName[num],this.properties.libraryFlag[num]);
  }

  private setFlag(num:number,flag:boolean): void {
    console.log("setFlag",num,flag);
    this.properties.flags[num] = flag;

    return;
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    //SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    //SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");

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
