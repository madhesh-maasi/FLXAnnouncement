import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './FlxannouncementsWebPart.module.scss';
import * as strings from 'FlxannouncementsWebPartStrings';
import "../../ExternalRef/Css/style.css";
import "../../ExternalRef/Css/Bootstrap.min.css";
import "../../ExternalRef/js/Bootstrap.js";
export interface IFlxannouncementsWebPartProps { 
  description: string;
}
  
export default class FlxannouncementsWebPart extends BaseClientSideWebPart<IFlxannouncementsWebPartProps> {

  public render(): void {   
    this.domElement.innerHTML = `
    <div class="cont">   
    <div class="row announcements-section justify-content-center">
    <div class="col-6 announcement border p-0">
    <h5 class="bg-secondary text-light px-4 py-2">Monthly Announcements</h5>
    <div class="add-announcements px-4 py-2 border-bottom"><a class="text-info">+ Add Announcements</a></div>
    <div class="announcement-list">  
    <ul class="list-unstyled"> 
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    </ul> 
    </div>
    </div>
    <!-- next Section-->   
    <div class="col-6 announcement border p-0">
    <h5 class="bg-secondary text-light px-4 py-2">Monthly Announcements</h5>
    <div class="add-announcements px-4 py-2 border-bottom"><a class="text-info">+ Add Announcements</a></div>
    <div class="announcement-list">  
    <ul class="list-unstyled"> 
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="#">FLX Announcements</a></li>
    </ul> 
    </div>
    </div>
    <!-- next Section-->
    <!-- next Section-->
    <div class="col-6 announcement border p-0">
    <h5 class="bg-secondary text-light px-4 py-2">Monthly Announcements</h5>
    <div class="add-announcements px-4 py-2 border-bottom"><a class="text-info">+ Add Announcements</a></div>
    <div class="announcement-list">  
    <ul class="list-unstyled"> 
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="#">FLX Announcements</a></li>
    </ul> 
    </div>
    </div>
    <!-- next Section-->      
    <!-- next Section-->
    <div class="col-6 announcement border p-0"> 
    <h5 class="bg-secondary text-light px-4 py-2">Monthly Announcements</h5>
    <div class="add-announcements px-4 py-2 border-bottom"><a class="text-info">+ Add Announcements</a></div>
    <div class="announcement-list">  
    <ul class="list-unstyled">  
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="#">FLX Announcements</a></li>
    </ul> 
    </div>
    </div>  
    <!-- next Section-->
    </div> 
    </div>
      `;  
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
