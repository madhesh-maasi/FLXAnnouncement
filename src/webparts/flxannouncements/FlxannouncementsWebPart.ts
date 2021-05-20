import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPComponentLoader } from "@microsoft/sp-loader";

SPComponentLoader.loadScript(
  // "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.4.min.js"
  "https://code.jquery.com/jquery-3.5.1.js"
);

import * as $ from "jquery";

import styles from './FlxannouncementsWebPart.module.scss';
import * as strings from 'FlxannouncementsWebPartStrings';
import { sp } from "@pnp/sp/presets/all";
import "../../ExternalRef/Css/style.css";
import "../../ExternalRef/Css/Bootstrap.min.css";
import "../../ExternalRef/js/Bootstrap.js";
import "../../ExternalRef/css/alertify.min.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");  

var siteURL = "";
var docurl ="";
var Filename=[];
var Fileupload=[];
var FileuploadEdit=[];
var allitems=[];
var editdata='';

export interface IFlxannouncementsWebPartProps { 
  description: string;
}
  
export default class FlxannouncementsWebPart extends BaseClientSideWebPart<IFlxannouncementsWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });  
    }); 
  } 

  public render(): void {   
    siteURL = this.context.pageContext.web.absoluteUrl;

    this.domElement.innerHTML = `
    <div class="cont"> 
     
    <div class="row announcements-section justify-content-center">
    <div class="col-6 announcement p-0">

    <div class="modal fade" id="announcementModal" tabindex="-1" aria-labelledby="announcementModalLabel" aria-hidden="true">
  <div class="modal-dialog announcement-modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="announcementModalLabel">Add Announcement</h5>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
      </div> 
      <div class="modal-body announcement-modal"> 
        <div class="row align-items-center my-3"><div class="col-4">Title</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="txttitle"></div></div>
        <div class="row align-items-center my-3"><div class="col-4">URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="txturl"></div></div>
        <div class="row align-items-center my-3"><div class="col-4">File</div><div class="col-1">:</div><div class="col-7"><input class="form-control-file custom-file-upload" type="file" id="uploadfile"></div></div>
        <div class="row align-items-center my-3"><div class="col-4">Document Type</div><div class="col-1">:</div><div class="col-7">
  
        <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">

  <input type="checkbox" class="btn-check" id="btnsensitive" autocomplete="off">
  <label class="btn btn-outline-theme" for="btnsensitive">Sensitive</label>
  
  <input type="checkbox" class="btn-check" id="btnvisible" autocomplete="off">
  <label class="btn btn-outline-theme" for="btnvisible">Visible</label>

  <input type="checkbox" class="btn-check" id="btnnewtab" autocomplete="off">
  <label class="btn btn-outline-theme" for="btnnewtab">Open a new tab</label>
</div>
 
        </div></div>
      </div>
      <div class="modal-footer"> 
        <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal" id="btnclose">Close</button>
        <button type="button" class="btn btn-sm btn-theme" id="btnsubmit">Submit</button> 
      </div>
    </div>
  </div>
</div>   

<div class="modal fade" id="announcementModalEdit" tabindex="-1" aria-labelledby="announcementModalLabel" aria-hidden="true">
<div class="modal-dialog announcement-modal-dialog">
  <div class="modal-content">
    <div class="modal-header">
      <h5 class="modal-title" id="announcementModalLabel">Edit Announcement</h5>
      <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
    </div> 
    <div class="modal-body announcement-modal"> 
      <div class="row align-items-center my-3"><div class="col-4">Title</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="edittitle"></div></div>
      <div class="row align-items-center my-3"><div class="col-4">Attachment URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="attachurl"></div></div>
      <div class="row align-items-center my-3"><div class="col-4">URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="editurl"></div></div>
      <div class="row align-items-center my-3"><div class="col-4">File</div><div class="col-1">:</div><div class="col-7"><input class="form-control-file custom-file-upload" type="file" id="uploadfileedit"></div></div>
      <div class="row align-items-center my-3"><div class="col-4">Document Type</div><div class="col-1">:</div><div class="col-7">

      <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">

<input type="checkbox" class="btn-check" id="editsensitive" autocomplete="off">
<label class="btn btn-outline-theme" for="editsensitive">Sensitive</label>

<input type="checkbox" class="btn-check" id="editvisible" autocomplete="off">
<label class="btn btn-outline-theme" for="editvisible">Visible</label>

<input type="checkbox" class="btn-check" id="editnewtab" autocomplete="off">
<label class="btn btn-outline-theme" for="editnewtab">Open a new tab</label>
</div>

      </div></div>
    </div>
    <div class="modal-footer"> 
      <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Close</button>
      <button type="button" class="btn btn-sm btn-theme" id="btnupdate">Update</button> 
    </div>
  </div>
</div>
</div> 

    <div class="border">
    <h5 class="bg-secondary text-light px-4 py-2">Monthly Announcements</h5>
    <div class="add-announcements px-4 py-2 border-bottom"><a class="text-info cursor" data-bs-toggle="modal" data-bs-target="#announcementModal">+ Add Announcements</a></div>
    <div class="announcement-list">  
    <ul class="list-unstyled" id="announcement-one"> 
    <!--<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>
    <li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="#">FLX Announcements</a></li>-->
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
      getFLXAnnouncement();
      $("#btnsubmit").click(async function()
      {
      await addItems();
      });
      $("#btnupdate").click(async function()
      {
      await updateItems();
      });
      $(document).on('click','.icon-edit',async function()
      {
      editdata=$(this).attr("data-id");
      console.log(editdata);
      
      $("#edittitle").val(allitems[editdata].Title);
      $("#attachurl").val(allitems[editdata].Url);
      //allitems[editdata].SensitiveDocument=true ? ($("#editsensitive").is(':checked')) : ($("#editsensitive").not(':checked')),
      //$("#editsensitive").val(allitems[editdata].SensitiveDocument = true ? ($("#editsensitive").is(':checked')) : ($("#editsensitive").not(':checked'))),
      //$("#editsensitive").attr(':checked');
      $("#editsensitive").prop( "checked", allitems[editdata].SensitiveDocument);
      $("#editvisible").prop("checked",allitems[editdata].Visible);
      $("#editnewtab").prop("checked",allitems[editdata].Openanewtab);
      
      });
      
      $("#btnclose").click(function()
      {
        $("#txttitle").val(""),
        $("#btnsensitive").val(""),
        $("#btnvisible").val(""),
        $("#btnnewtab").val(""), 
        $("#uploadfile").val(""),
        $("#txturl").val("")
      });
      $(document).on("change", "#uploadfile", function () {
        if ($(this)[0].files.length > 0) {
          for (let index = 0; index < $(this)[0].files.length; index++) {
            const file = $("#uploadfile")[0]["files"][index];
            Fileupload.push(file);
          }
          //$(this).val("");
          $(this).parent().find("label").text("Choose File");
        }
      });
      $(document).on("change", "#uploadfileedit", function () {
        if ($(this)[0].files.length > 0) {
          for (let index = 0; index < $(this)[0].files.length; index++) {
            const file = $("#uploadfileedit")[0]["files"][index];
            FileuploadEdit.push(file);
          }
          //$(this).val("");
          $(this).parent().find("label").text("Choose File");
        }
      });
     
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

async function getFLXAnnouncement()
{
  await sp.web.lists.getByTitle("FLXAnnouncement").items.select("*").filter("Visible eq '" + 1 + "'").get().then(async (item)=>
  {
var htmlforannouncement="";
allitems=item;
console.log(allitems);

for(var i=0;i<item.length;i++){
  Filename.push(item[i].Url.split('/').pop());
  console.log("Filename");
console.log(Filename);

  if(item[i].SensitiveDocument==true){
  if(item[i].Openanewtab==true){
    if (Filename[i].split(".").pop() == "pdf")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="${item[i].Url}" target="_blank" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "ppt")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="${item[i].Url}" target="_blank" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="${item[i].Url}" target="_blank" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-excel"></span><a href="${item[i].Url}" target="_blank" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="${item[i].Url}" target="_blank" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-new"></span><a href="${item[i].Url}" target="_blank" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }

}

else {
  if (Filename[i].split(".").pop() == "pdf")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="${item[i].Url}" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "ppt")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="${item[i].Url}" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="${item[i].Url}" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-excel"></span><a href="${item[i].Url}" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="${item[i].Url}" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-new"></span><a href="${item[i].Url}" onclick="return confirm('Are you sure?')">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
}
}

else{
  if(item[i].Openanewtab==true){
    if (Filename[i].split(".").pop() == "pdf")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="${item[i].Url}" target="_blank">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "ppt")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="${item[i].Url}" target="_blank">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="${item[i].Url}" target="_blank">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-excel"></span><a href="${item[i].Url}" target="_blank">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="${item[i].Url}" target="_blank">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-new"></span><a href="${item[i].Url}" target="_blank">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }

}

else {
  if (Filename[i].split(".").pop() == "pdf")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-pdf"></span><a href="${item[i].Url}">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "ppt")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-ppt"></span><a href="${item[i].Url}">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-doc"></span><a href="${item[i].Url}">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-excel"></span><a href="${item[i].Url}">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-img"></span><a href="${item[i].Url}">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
    else
    {
  htmlforannouncement+=`<li class="py-2 px-4 d-flex align-items-center"><span class="announce-icon announce-new"></span><a href="${item[i].Url}">${Filename[i]}</a><span class="icon-edit" data-id=${i} data-bs-toggle="modal" data-bs-target="#announcementModalEdit"></span>`;
    }
}
}

}
$("#announcement-one").html("");
$("#announcement-one").html(htmlforannouncement);
  }).catch((error)=>
  {
    console.log(error);
  });
  }

  async function addItems() {
    var requestdata = {}; 
     if (Fileupload.length > 0) {
      await Fileupload.map((filedata) => {
            sp.web
              .getFolderByServerRelativeUrl("/sites/FLXCommunity/AnnouncementDocument/")
              .files.add(filedata.name, filedata, true)
              .then (function(data){
                console.log(data);
                requestdata = {
                  Title: $("#txttitle").val(),
                  // DocumentUrl:{
                  //   "__metadata": { type: "SP.FieldUrlValue" },
                  //   Description: "",
                  //   Url: $("#txturl").val(),
                
                  // },
                  Url:data.data.ServerRelativeUrl,
                  SensitiveDocument: $("#btnsensitive").is(':checked') ? true : false,
                  Visible: $("#btnvisible").is(':checked') ? true : false,
                  Openanewtab: $("#btnnewtab").is(':checked') ? true : false,
                };
                sp.web.lists
                .getByTitle("FLXAnnouncement")
                .items.add(requestdata)
                .then(function (data) {
                  console.log(data);
                  AlertMessage("Submitted successfully");
                })
                .catch(function (error) {
                  ErrorCallBack(error, "addItems");
                });
              })
          });
    }  
    else{
      requestdata = {
        Title: $("#txttitle").val(),
        // DocumentUrl:{
        //   "__metadata": { type: "SP.FieldUrlValue" },
        //   Description: "",
        //   Url: $("#txturl").val(),
      
        // },
        Url:$("#txturl").val(),
        SensitiveDocument: $("#btnsensitive").is(':checked') ? true : false,
        Visible: $("#btnvisible").is(':checked') ? true : false,
        Openanewtab: $("#btnnewtab").is(':checked') ? true : false,
      };
      sp.web.lists
      .getByTitle("FLXAnnouncement")
      .items.add(requestdata)
      .then(function (data) {
        console.log(data);
        AlertMessage("Submitted successfully");
      })
      .catch(function (error) {
        ErrorCallBack(error, "addItems");
      });
    }
  }

  async function updateItems() {

    var requestdata = {}; 
    var Id=allitems[editdata].ID;
     if (FileuploadEdit.length > 0) {
      await FileuploadEdit.map((filedata) => {
            sp.web
              .getFolderByServerRelativeUrl("/sites/FLXCommunity/AnnouncementDocument/")
              .files.add(filedata.name, filedata, true)
              .then (function(data){
                console.log(data);
                requestdata = {
                  Title: $("#edittitle").val(),
                  // DocumentUrl:{
                  //   "__metadata": { type: "SP.FieldUrlValue" },
                  //   Description: "",
                  //   Url: $("#txturl").val(),
                
                  // },
                  Url:data.data.ServerRelativeUrl,
                  SensitiveDocument: $("#editsensitive").is(':checked') ? true : false,
                  Visible: $("#editvisible").is(':checked') ? true : false,
                  Openanewtab: $("#editnewtab").is(':checked') ? true : false,
                };
                sp.web.lists
                .getByTitle("FLXAnnouncement")
                .items.getById(Id).update(requestdata)
                .then(function (data) {
                  console.log(data);
                  AlertMessage("Updated successfully");
                })
                .catch(function (error) {
                  ErrorCallBack(error, "updateItems");
                });
              })
          });
    } 
    else{
      requestdata = {
        Title: $("#edittitle").val(),
        Url:$("#editurl").val(),
        SensitiveDocument: $("#editsensitive").is(':checked') ? true : false,
        Visible: $("#editvisible").is(':checked') ? true : false,
        Openanewtab: $("#editnewtab").is(':checked') ? true : false,
      };
      sp.web.lists
      .getByTitle("FLXAnnouncement")
      .items.getById(Id).update(requestdata)
      .then(function (data) {
        console.log(data);
        AlertMessage("Updated successfully");
      })
      .catch(function (error) {
        ErrorCallBack(error, "updateItems");
      });

    } 
  }

  async function ErrorCallBack(error, methodname) 
  {
    try {
      var errordata = {
        Error: error.message,
        MethodName: methodname,
      };
      await sp.web.lists
        .getByTitle("ErrorLog")
        .items.add(errordata)
        .then(function (data) 
        {
          $('.loader').hide();
          AlertMessage("Something went wrong.please contact system admin");
        });
    } catch (e) {
      //alert(e.message);
      $('.loader').hide();
      Alert("Something went wrong.please contact system admin");
    }
  }
  function AlertMessage(strMewssageEN) {
    alertify
      .alert()
      .setting({
        label: "OK",
        
        message: strMewssageEN,
  
        onok: function () {
          window.location.href = "#";
          location.reload();
        },
      })
      
      .show()
      .setHeader("<em>Confirmation</em> ")
      .set("closable", false);
  }
  
  function Alert(strMewssageEN) {
    alertify
      .alert()
      .setting({
        label: "OK",
        
        message: strMewssageEN,
  
        onok: function () {
          window.location.href = "#";
        },
      })
      
      .show()
      .setHeader("<em>Confirmation</em> ")
      .set("closable", false);
  }