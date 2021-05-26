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

import styles from './FlxwhitepapersWebPart.module.scss';
import * as strings from 'FlxwhitepapersWebPartStrings';
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
var urlFile = "";
var updateUrlFile = "";

export interface IFlxwhitepapersWebPartProps {
  description: string;
}
let SelectedFileName = ""
export default class FlxwhitepapersWebPart extends BaseClientSideWebPart<IFlxwhitepapersWebPartProps> {
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
     
    <div class="row announcements-section">
    <div class="col-6 announcement p-0">

    <div class="modal fade" id="whiteannouncementModal" tabindex="-1" aria-labelledby="announcementModalLabel" aria-hidden="true">
  <div class="modal-dialog announcement-modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="whiteannouncementModalLabel">Add White Papers Announcement</h5>
      </div> 
      <div class="modal-body announcement-modal"> 
        <div class="row align-items-center my-3"><div class="col-4">Title</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="whitetxttitle"></div></div>
        
        
        <div class="row align-items-center my-3"><div class="col-4">Source</div><div class="col-1">:</div><div class="col-7 clsRadioSec">
        
        <label><input type="radio" class="radioc" name="urlFile" id="whiteurlRadio" value="Url"> Url </label>
        <label><input type="radio"  class="radioc" name="urlFile" id="whitefileRadio" value="File"> File</label>


        
        </div></div>


        <div class="row align-items-center my-3 radioToggle" id="whiteurlSection" style="display:none"><div class="col-4">URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="whitetxturl"></div></div>
        <div class="row align-items-center my-3 radioToggle" id="whitefileSection" style="display:none"><div class="col-4">File</div><div class="col-1">:</div><div class="col-7"><input class="form-control-file custom-file-upload" type="file" id="whiteuploadfile"></div></div>
        <div class="row align-items-center my-3"><div class="col-4">Document Type</div><div class="col-1">:</div><div class="col-7">
  
        <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">

  <input type="checkbox" class="btn-check" id="whitebtnsensitive" autocomplete="off">
  <label class="btn btn-outline-theme" for="whitebtnsensitive">Sensitive</label>
  
  <input type="checkbox" class="btn-check" id="whitebtnvisible" autocomplete="off">
  <label class="btn btn-outline-theme" for="whitebtnvisible">Visible</label>

  <input type="checkbox" class="btn-check" id="whitebtnnewtab" autocomplete="off">
  <label class="btn btn-outline-theme" for="whitebtnnewtab">Open a new tab</label>
</div>
 
        </div></div>
      </div>
      <div class="modal-footer"> 
        <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal" id="whitebtnclose">Close</button>
        <button type="button" class="btn btn-sm btn-theme" id="whitebtnsubmit">Submit</button> 
      </div>
    </div>
  </div>
</div>   

<div class="modal fade" id="whiteannouncementModalEdit" tabindex="-1" aria-labelledby="announcementModalLabel" aria-hidden="true">
<div class="modal-dialog announcement-modal-dialog">
  <div class="modal-content">
    <div class="modal-header">
      <h5 class="modal-title" id="whiteannouncementModalLabel">Edit White Papers Announcement</h5>
      <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
    </div>  
    <div class="modal-body announcement-modal"> 
      <div class="row align-items-center my-3"><div class="col-4">Title</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="whiteedittitle"></div></div>
      <!--<div class="row align-items-center my-3"><div class="col-4">Source</div><div class="col-1">:</div><div class="col-7 clsRadioSec">
        <label><input type="radio" class="Eradioc" name="EurlFile" id="whiteEurlRadio" value="Url"> Url </label>
        <label><input type="radio"  class="Eradioc" name="EurlFile" id="whiteEfileRadio" value="File"> File</label>
        </div></div>-->
      <div class="row align-items-center my-3" id="whiteEurlSection" style="display:none"><div class="col-4">URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="whiteediturl"></div></div>
      <div class="row align-items-start my-3" id="whiteEfileSection" style="display:none"><div class="col-4">File</div><div class="col-1">:</div><div class="col-7" id="whiteeditFUploadSec"><div><input class="form-control-file custom-file-upload" type="file" id="whiteuploadfileedit"></div><div class="uploadedFile mt-1"></div></div></div>
      <div class="row align-items-center my-3"><div class="col-4">Document Type</div><div class="col-1">:</div><div class="col-7">
  
      <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">

<input type="checkbox" class="btn-check" id="whiteeditsensitive" autocomplete="off">
<label class="btn btn-outline-theme" for="whiteeditsensitive">Sensitive</label>

<input type="checkbox" class="btn-check" id="whiteeditvisible" autocomplete="off">
<label class="btn btn-outline-theme" for="whiteeditvisible">Visible</label>

<input type="checkbox" class="btn-check" id="whiteeditnewtab" autocomplete="off">
<label class="btn btn-outline-theme" for="whiteeditnewtab">Open a new tab</label>
</div>

      </div></div>
    </div>
    <div class="modal-footer justify-content-between"> 
    <div>
    <button type="button" class="btn btn-sm btn-danger" id="whiteAnABtnDelete" data-bs-toggle="modal" data-bs-target="#whiteAnADeleteModal">Delete</button>
     </div>
      <div class="d-flex">
      <button type="button" class="btn btn-sm btn-secondary mx-1"  id = "whitebtnUpdateClose" data-bs-dismiss="modal">Close</button>
      <button type="button" class="btn btn-sm btn-theme mx-1" id="whitebtnupdate">Update</button> </div>
    </div>
  </div>
</div>
</div> 

<div class="modal fade" id="whiteAnADeleteModal" tabindex="-1" aria-labelledby="AnADeleteModalLabel" aria-hidden="true">
  <div class="modal-dialog AnA-delete-warning-dialog">
    <div class="modal-content">
      <div class="modal-header">
        
      </div>
      <div class="modal-body AnA-delete-warning text-center pt-5"> 
      <h5 class="modal-title" id="whitedeleteAlterModalLabel">Confirmation</h5>
      <p class="mb-0">Are you sure want to Delete?</p>
      </div>
      <div class="modal-footer">
        <button type="button" id="whitecancelAnADelete" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">No</button>
        <button type="button" id="whiteconfirmAnADelete" class="btn btn-sm btn-danger">Yes</button>
      </div>
    </div>
  </div>
</div>

    <div class="border announcement-sec">
    <h5 class="bg-secondary text-light px-4 py-2" id="whiteheaderTitle">White Papers We Announce</h5>
    <div class="add-announcements px-4 py-2 border-bottom"><a class="text-info cursor" data-bs-toggle="modal" data-bs-target="#whiteannouncementModal">+ Add White Papers</a></div>
    <div id="whiteannouncement-list">  
    <ul class="list-unstyled" id="whiteannouncement-one"> 
    
    </ul> 
    </div>  
    </div>
    </div>
   
    </div> 
    </div>
    `;
    //$("#whiteheaderTitle").text(headerTitle)
    $("#whiteAnABtnDelete").click(()=>{
      $(".announcement-modal-dialog").hide();
    });
    $("#whitecancelAnADelete").click(()=>{
      $(".announcement-modal-dialog").show();
    });
    $("input[type=radio][name=urlFile]").change(function(e) {
      urlFile = e.currentTarget.value;
      console.log(urlFile);
      if(e.currentTarget.value == "Url"){
        $("#whiteurlSection").show();
        $("#whitefileSection").hide();
      }else if (e.currentTarget.value == "File"){
        $("#whiteurlSection").hide();
        $("#whitefileSection").show();
        $("#whitetxturl").val("");
      }
          });
    $("input[type=radio][name=EurlFile]").change(function(e) {
            updateUrlFile = e.currentTarget.value;
            console.log(updateUrlFile);
            if(e.currentTarget.value == "Url"){
              $("#whiteEurlSection").show();
              $("#whiteEfileSection").hide();
            }else if (e.currentTarget.value == "File"){
              $("#whiteEurlSection").hide();
              $("#whiteEfileSection").show();
              $("#whiteediturl").val("");
            }
                });

    getFLXWhitePaperAnnouncements();
    $(document).on('click','.sensitivewhite',async function(e)
      {
        let result=confirm("Are you sure want to proceed?");
        if(result==true){
        return result;
      }
      else{
        e.stopImmediatePropagation();
        e.preventDefault();
      }
      });
    $("#whitebtnsubmit").click(async function()
    {
    await addItems();
    });
    $("#whitebtnupdate").click(async function()
    {
    await updateItems();
    });
    $(document).on('click','.icon-edit-announce',async function()
    {
        // FileFormFolder
    editdata=$(this).attr("data-id"); 
    console.log(editdata); 
    
    $("#whiteedittitle").val(allitems[editdata].Title);
    $("#whiteediturl").val(allitems[editdata].Url); 
 
    if(allitems[editdata].UrlOrFile == "File"){
      SelectedFileName = allitems[editdata].Url.split('/').pop();
      $(".uploadedFile").html("");
      $(".uploadedFile").html(`<a href="${allitems[editdata].Url}">${SelectedFileName}</a>`);
    }else{
      $(".uploadedFile").html("")
    }
       
    if(allitems[editdata].UrlOrFile == "Url"){
      updateUrlFile = "Url"
      $("#whiteEurlRadio").prop("checked",true);
      $("#whiteEurlSection").show();
      $("#whiteEfileSection").hide();
    } else{
      updateUrlFile = "File"
      $("#whiteEfileRadio").prop("checked",true);
      $("#whiteEurlSection").hide();
      $("#whiteEfileSection").show();
      // $("#whiteeditFUploadSec").html(`<input class="form-control-file custom-file-upload" type="file" id="whiteuploadfileedit">`)
    }
    console.log(`InList : ${urlFile}`);
    
    $("#whiteeditsensitive").prop( "checked", allitems[editdata].SensitiveDocument);
    $("#whiteeditvisible").prop("checked",allitems[editdata].Visible);
    $("#whiteeditnewtab").prop("checked",allitems[editdata].Openanewtab);
    
    });
    
    $("#whitebtnclose").click(function()
    {
      $("#whitetxttitle").val("");
      $("#whitebtnsensitive").val("");
      $("#whitebtnvisible").val("");
      $("#whitebtnnewtab").val("");
      $("#whiteuploadfile").val("");
      $("#whitetxturl").val("");
      
      let radioReset = document.getElementsByName("urlFile");
      for(var i=0;i<radioReset.length;i++)
      radioReset[i]["checked"] = false;
    });
    $("#whitebtnUpdateClose").click(()=>{
      $("#whiteuploadfileedit").val("")

    })
    $(document).on("change", "#whiteuploadfile", function () {
      if ($(this)[0].files.length > 0) {
        for (let index = 0; index < $(this)[0].files.length; index++) {
          const file = $("#whiteuploadfile")[0]["files"][index];
          Fileupload.push(file);
        }
        //$(this).val("");
        $(this).parent().find("label").text("Choose File");
      }
    });
    $(document).on("change", "#whiteuploadfileedit", function () {
      
      if ($(this)[0].files.length > 0) {
        for (let index = 0; index < $(this)[0].files.length; index++) {
          const file = $("#whiteuploadfileedit")[0]["files"][index];
          FileuploadEdit.push(file);
        }
        //$(this).val("");
        $(this).parent().find("label").text("Choose File");
        $(".uploadedFile").html("")
      }else{
        $(".uploadedFile").html(`<a href="${allitems[editdata].Url}">${SelectedFileName}</a>`);
      }
    });
    $("#whiteconfirmAnADelete").click(()=>{
      deleteAnA(allitems[editdata].ID)
    })
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

async function getFLXWhitePaperAnnouncements()
{


  // if(FLXWhitePaperAnnouncements)
  // {
    await sp.web.lists.getByTitle("FLXWhitePaperAnnouncements").items.select("*").filter("Visible eq '" + 1 + "'").get().then(async (item)=>
    {
  var htmlforwhiteannouncement="";
  allitems=item;
  console.log(allitems);
  if(item.length  == 0){
    
    $("#whiteannouncement-list").html(`<div class="text-center pt-5">No Items Available</div>`)
  }
  for(var i=0;i<item.length;i++){
    Filename.push(item[i].Url.split('/').pop());
    console.log("Filename");
  console.log(Filename);
  
    if(item[i].SensitiveDocument==true){
    if(item[i].Openanewtab==true){
      if (Filename[i].split(".").pop() == "pdf")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
  
  }
  
  else {
    if (Filename[i].split(".").pop() == "pdf")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf col-2"></span><a href="${item[i].Url}" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt col-2"></span><a href="${item[i].Url}" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc col-2"></span><a href="${item[i].Url}" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel col-2"></span><a href="${item[i].Url}" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img col-2"></span><a href="${item[i].Url}" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new col-2"></span><a href="${item[i].Url}" class="col-8 sensitivewhite">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
  }
  }
  
  else{
    if(item[i].Openanewtab==true){ 
      if (Filename[i].split(".").pop() == "pdf")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new col-2"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
  
  }
  
  else {
    if (Filename[i].split(".").pop() == "pdf")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf col-2"></span><a href="${item[i].Url}" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt col-2"></span><a href="${item[i].Url}" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc col-2"></span><a href="${item[i].Url}" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel col-2"></span><a href="${item[i].Url}" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img col-2"></span><a href="${item[i].Url}" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlforwhiteannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new col-2"></span><a href="${item[i].Url}" class="col-8">${Filename[i]}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#whiteannouncementModalEdit"></div></li>`;
      }
  } 
  }
   
  }
  $("#whiteannouncement-one").html("");
  $("#whiteannouncement-one").html(htmlforwhiteannouncement);
    }).catch((error)=>
    {
      console.log(error);
    });
  }
  // else{
  //   $("#announcement-one").html("");
  // $("#announcement-one").html(`<li class="py-2 px-4 d-flex align-items-center row">No data to display or Please select list name</li>`);
  // }
  // }

  

  async function addItems() {
    var requestdata = {}; 
     if (Fileupload.length > 0) {
      await Fileupload.map((filedata) => {
            sp.web
              .getFolderByServerRelativeUrl("/sites/FLXCommunity/AnnouncementDocument/FLXWhitePaperAnnouncements")
              .files.add(filedata.name, filedata, true)
              .then (function(data){
                console.log(data);
                requestdata = {
                  Title: $("#whitetxttitle").val(),
                  // DocumentUrl:{
                  //   "__metadata": { type: "SP.FieldUrlValue" },
                  //   Description: "",
                  //   Url: $("#txturl").val(),
                
                  // },
                  Url:data.data.ServerRelativeUrl,
                  SensitiveDocument: $("#whitebtnsensitive").is(':checked') ? true : false,
                  Visible: $("#whitebtnvisible").is(':checked') ? true : false,
                  Openanewtab: $("#whitebtnnewtab").is(':checked') ? true : false,
                  UrlOrFile:urlFile
                };
                sp.web.lists
                .getByTitle("FLXWhitePaperAnnouncements")
                .items.add(requestdata)
                .then(function (data) {
                  console.log(data);
                  AlertMessage("<div class='alertfy-success'>Submitted successfully</div>");
                })
                .catch(function (error) {
                  // ErrorCallBack(error, "addItems");
                  console.log(error);
                  
                });
              })
          });
    }  
    else{
      requestdata = {
        Title: $("#whitetxttitle").val(),
        // DocumentUrl:{
        //   "__metadata": { type: "SP.FieldUrlValue" },
        //   Description: "",
        //   Url: $("#txturl").val(),
      
        // },
        Url:$("#whitetxturl").val(),
        SensitiveDocument: $("#whitebtnsensitive").is(':checked') ? true : false,
        Visible: $("#whitebtnvisible").is(':checked') ? true : false,
        Openanewtab: $("#whitebtnnewtab").is(':checked') ? true : false,
        UrlOrFile:urlFile
      };
      sp.web.lists
      .getByTitle("FLXWhitePaperAnnouncements")
      .items.add(requestdata)
      .then(function (data) {
        console.log(data);
        AlertMessage("<div class='alertfy-success'>Submitted successfully</div>");
      })
      .catch(function (error) {
        // ErrorCallBack(error, "addItems");
        console.log(error);
        
      });
    }
  }

  async function updateItems() {
console.log(FileuploadEdit);

    var requestdata = {}; 
    var Id=allitems[editdata].ID;
     if (FileuploadEdit.length > 0) {
      await FileuploadEdit.map((filedata) => {
            sp.web
              .getFolderByServerRelativeUrl("/sites/FLXCommunity/AnnouncementDocument/FLXWhitePaperAnnouncements")
              .files.add(filedata.name, filedata, true)
              .then (function(data){
                console.log(data);
                requestdata = {
                  Title: $("#whiteedittitle").val(),
                  // DocumentUrl:{
                  //   "__metadata": { type: "SP.FieldUrlValue" },
                  //   Description: "",
                  //   Url: $("#txturl").val(),
                  // },
                  Url:data.data.ServerRelativeUrl,
                  SensitiveDocument: $("#whiteeditsensitive").is(':checked') ? true : false,
                  Visible: $("#whiteeditvisible").is(':checked') ? true : false,
                  Openanewtab: $("#whiteeditnewtab").is(':checked') ? true : false,
                  UrlOrFile:updateUrlFile
                };
                sp.web.lists
                .getByTitle("FLXAnnouncement")
                .items.getById(Id).update(requestdata)
                .then(function (data) {
                  console.log(data);
                  AlertMessage("<div class='alertfy-success'>Updated successfully</div>");
                })
                .catch(function (error) {
                  // ErrorCallBack(error, "updateItems");
                  console.log(error);
                  
                });
              })
          }); 
    } else if(FileuploadEdit.length == 0 && updateUrlFile == "File" && SelectedFileName != ""){
      requestdata = {
        Title: $("#whiteedittitle").val(),
        // DocumentUrl:{
        //   "__metadata": { type: "SP.FieldUrlValue" },
        //   Description: "",
        //   Url: $("#txturl").val(),
        // },
        
        SensitiveDocument: $("#whiteeditsensitive").is(':checked') ? true : false,
        Visible: $("#whiteeditvisible").is(':checked') ? true : false,
        Openanewtab: $("#whiteeditnewtab").is(':checked') ? true : false,
        UrlOrFile:updateUrlFile
      };
      sp.web.lists
      .getByTitle("FLXWhitePaperAnnouncements")
      .items.getById(Id).update(requestdata)
      .then(function (data) {
        console.log(data);
        AlertMessage("<div class='alertfy-success'>Updated successfully</div>");
      })
      .catch(function (error) {
        // ErrorCallBack(error, "updateItems");
        console.log(error);
        
      });
    }else if(FileuploadEdit.length == 0 && updateUrlFile == "File" && SelectedFileName == ""){
      $(".uploadedFile").html(`<p class="text-danger">File Cannot be Empty</p>`)
    }
    else{
      requestdata = {
        Title: $("#whiteedittitle").val(),
        Url:$("#whiteediturl").val(),
        SensitiveDocument: $("#whiteeditsensitive").is(':checked') ? true : false,
        Visible: $("#whiteeditvisible").is(':checked') ? true : false,
        Openanewtab: $("#whiteeditnewtab").is(':checked') ? true : false,
        UrlOrFile:updateUrlFile
      };
      sp.web.lists
      .getByTitle("FLXWhitePaperAnnouncements")
      .items.getById(Id).update(requestdata)
      .then(function (data) {
        console.log(data);
        AlertMessage("<div class='alertfy-success'>Updated successfully</div>");
      })
      .catch(function (error) {
        // ErrorCallBack(error, "updateItems");
        console.log(error);
        
      }); 

    } 
  }
 
  // async function ErrorCallBack(error, methodname) 
  // {
  //   try {
  //     var errordata = {
  //       Error: error.message,
  //       MethodName: methodname,
  //     };
  //     await sp.web.lists
  //       .getByTitle("ErrorLog")
  //       .items.add(errordata)
  //       .then(function (data) 
  //       {
  //         $('.loader').hide();
  //         AlertMessage("Something went wrong.please contact system admin");
  //       });
  //   } catch (e) {
  //     //alert(e.message);
  //     $('.loader').hide();
  //     Alert("Something went wrong.please contact system admin");
  //   }
  // }
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
      .setHeader("<div class='fw-bold alertifyConfirmation'>Confirmation</div> ")
      .set("closable", false);
  }
const deleteAnA = (id) =>{
   sp.web.lists.getByTitle("FLXWhitePaperAnnouncements").items.getById(id).delete().then(()=>{location.reload()}).catch((error)=>{alert("Error Occured");})
}