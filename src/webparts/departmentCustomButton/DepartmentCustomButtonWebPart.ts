import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DepartmentCustomButtonWebPart.module.scss';
import * as strings from 'DepartmentCustomButtonWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IDepartmentCustomButtonWebPartProps {
  ButtonType1: string;
}
export interface Button2Lists {
  value: Button2Lists[];
}

export interface Button2Lists {
  Title: string;
  Id: string;
  BackgroundColor_x0028_Hex_x0029_: string;
  TextColor_x0028_Hex_x0029_: string;
  Url:string;
  BorderColor_x0028_Hex_x0029_:string;
  Link_x0028_Url_x0029_:{Url:string};
}

export interface logo {
  value: logo[];
}

export interface logo {
  ServerRelativeUrl:string;
 
}


export default class DepartmentCustomButtonWebPart extends BaseClientSideWebPart<IDepartmentCustomButtonWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <div>
     <section class="${styles.helloWord} ${
       !!this.context.sdks.microsoftTeams ? styles.teams : ""
     }" data-automation-id="CustomButtons">     
       <div id="splogoListContainer" style="display: flex;
       flex-direction: row-reverse;"/>
       </div>
       <div id="spButton2ListContainer" class="${styles.headerButton1}" /></div>
     </section>
    </div>
    `;
     this._renderButton2ListAsync();
     this._renderTextLogoListAsync();
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

    // button 
    private _getListButton2Data(): Promise<Button2Lists> {
      return this.context.spHttpClient
        .get(
          `https://rspsgp.sharepoint.com/sites/Intranet/_api/web/lists/getbytitle('Button%20Settings%20(Department%20Sites)')/items`,
          SPHttpClient.configurations.v1
        )
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    }
  
    private _renderButton2ListAsync(): void {
      this._getListButton2Data().then((response) => {
        this._renderButton2List(response.value);
      });
    }
  
    private _renderButton2List(items: Button2Lists[]): void {
      let html: string = "";
      let count: number = 4;
      items.forEach((item1: Button2Lists, index) => {
  
        if (count === index + 1) {
          html += `
        <button type="button" onclick="location.href='${item1.Link_x0028_Url_x0029_.Url}';" class="${styles.button2}" style="color:${item1.TextColor_x0028_Hex_x0029_};background-color:${item1.BackgroundColor_x0028_Hex_x0029_}; border-color:${item1.BorderColor_x0028_Hex_x0029_};font-size:${escape(this.properties.ButtonType1)}px !important;"> ${item1.Title}</button>
        `;
          count = count + 3;
        } else {
          html += `
          <button type="button" onclick="location.href='${item1.Link_x0028_Url_x0029_.Url}';" class="${styles.button2}" style="color:${item1.TextColor_x0028_Hex_x0029_};background-color:${item1.BackgroundColor_x0028_Hex_x0029_};border-color:${item1.BorderColor_x0028_Hex_x0029_}; font-size:${escape(this.properties.ButtonType1)}px !important;"> ${item1.Title}</button>
          `;
        } 
      });
  
      const listContainer: Element = this.domElement.querySelector(
        "#spButton2ListContainer"
      );
      listContainer.innerHTML = html;
    }

      //text and logo
  private _getListTextLogoData(): Promise<logo> {
    return this.context.spHttpClient
      .get(
       `https://rspsgp.sharepoint.com/sites/Intranet/_api/web/GetFolderByServerRelativeUrl('logo%20image')/Files`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderTextLogoListAsync(): void {
    this._getListTextLogoData().then((response) => {
      this._renderTextLogoList(response.value);
    });
  }

  private _renderTextLogoList(items: logo[]): void {
    let html: string = "";

    for(let i= 0; i<1;i++)
    {
      html += `
      <a href="https://rspsgp.sharepoint.com/sites/Intranet" ><img alt="" src="${
        items[i].ServerRelativeUrl
      }" class="${styles.logo}" /></a>
          `;
    }
    const listContainer2: Element = this.domElement.querySelector(
      "#splogoListContainer"
    );
    listContainer2.innerHTML = html;
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
                PropertyPaneTextField("ButtonType1", {
                  label:"Button Group Font Size (px)",
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
