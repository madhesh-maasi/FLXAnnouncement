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
    <div class="col-6 announcement p-0">

    <div class="announcement-btn-secton my-2 text-end">
    <button class="btn btn-theme" data-bs-toggle="modal" data-bs-target="#announcementModal">Add</button>
    </div>
    <div class="modal fade" id="announcementModal" tabindex="-1" aria-labelledby="announcementModalLabel" aria-hidden="true">
  <div class="modal-dialog announcement-modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="announcementModalLabel">Add Announcement</h5>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
      </div> 
      <div class="modal-body announcement-modal"> 
        <div class="row align-items-center my-3"><div class="col-4">Title</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text"></div></div>
        <div class="row align-items-center my-3"><div class="col-4">URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text"></div></div>
        <div class="row align-items-center my-3"><div class="col-4">File</div><div class="col-1">:</div><div class="col-7"><input class="form-control-file custom-file-upload" type="file"></div></div>
        <div class="row align-items-center my-3"><div class="col-4">Options</div><div class="col-1">:</div><div class="col-7">
  
        <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">
  <input type="checkbox" class="btn-check" id="btncheck1" autocomplete="off">
  <label class="btn btn-outline-theme" for="btncheck1">Checkbox 1</label>
  
  <input type="checkbox" class="btn-check" id="btncheck2" autocomplete="off">
  <label class="btn btn-outline-theme" for="btncheck2">Checkbox 2</label>

  <input type="checkbox" class="btn-check" id="btncheck3" autocomplete="off">
  <label class="btn btn-outline-theme" for="btncheck3">Checkbox 3</label>
</div>
 
        </div></div>
      </div>
      <div class="modal-footer"> 
        <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Close</button>
        <button type="button" class="btn btn-sm btn-theme">Submit</button> 
      </div>
    </div>
  </div>
</div>   
    <div class="border">
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
