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

import styles from './FlxnewsletterWebPart.module.scss';
import * as strings from 'FlxnewsletterWebPartStrings';
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
var Sdata="";
export interface IFlxnewsletterWebPartProps {
  description: string;
}
let SelectedFileName = ""
export default class FlxnewsletterWebPart extends BaseClientSideWebPart<IFlxnewsletterWebPartProps> {
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
    <div class="loader-section" style="display:none"> 
    <div class="loader"></div>  
    </div></div> 
    <div class="cont"> 
     
    <div class="row announcements-section">
    <div class="col-6 announcement p-0">

    <div class="modal fade" id="newsannouncementModal" tabindex="-1" aria-labelledby="announcementModalLabel" aria-hidden="true">
  <div class="modal-dialog announcement-modal-dialog">
    <div class="modal-content rounded-0">
      <div class="modal-header">
        <h5 class="modal-title fw-bold w-100 text-center" id="newsannouncementModalLabel">Add Newsletter Announcement</h5>
      </div>  
      <div class="modal-body announcement-modal"> 
        <div class="row align-items-center my-3"><div class="col-4">Title</div><div class="col-1">:</div><div class="col-7"><input class="form-control rounded-0" type="text" id="newstxttitle"></div></div>
        
        
        <div class="row align-items-center my-3"><div class="col-4">Source</div><div class="col-1">:</div><div class="col-7 clsRadioSec">
        
        <label><input type="radio" class="radioc" name="urlFile" id="newsurlRadio" value="Url"> Url </label>
        <label><input type="radio"  class="radioc" name="urlFile" id="newsfileRadio" value="File"> File</label>


        
        </div></div>


        <div class="row align-items-center my-3 radioToggle" id="newsurlSection" style="display:none"><div class="col-4">URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control rounded-0" type="text" id="newstxturl"></div></div>
        <div class="row align-items-center my-3 radioToggle" id="newsfileSection" style="display:none"><div class="col-4">File</div><div class="col-1">:</div><div class="col-7"><input class="form-control-file custom-file-upload" type="file" id="newsuploadfile"></div></div>
        <div class="row align-items-center my-3"><div class="col-4">Document/Url Properties</div><div class="col-1">:</div><div class="col-7">
  
        <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">

  <input type="checkbox" class="btn-check" id="newsbtnsensitive" autocomplete="off">
  <label class="btn btn-outline-theme" for="newsbtnsensitive">Sensitive</label>
  
  <input type="checkbox" class="btn-check" id="newsbtnvisible" autocomplete="off">
  <label class="btn btn-outline-theme" for="newsbtnvisible">Visible</label>

  <input type="checkbox" class="btn-check" id="newsbtnnewtab" autocomplete="off">
  <label class="btn btn-outline-theme" for="newsbtnnewtab">Open a new tab</label>
</div>
 
        </div></div>
      </div>
      <div class="modal-footer"> 
        <button type="button" class="btn btn-sm btn-secondary rounded-0" data-bs-dismiss="modal" id="newsbtnclose">Close</button>
        <button type="button" class="btn btn-sm btn-theme rounded-0" id="newsbtnsubmit">Submit</button> 
      </div>
    </div>
  </div>
</div>   

<div class="modal fade" id="newsannouncementModalEdit" tabindex="-1" aria-labelledby="announcementModalLabel" aria-hidden="true">
<div class="modal-dialog announcement-modal-dialog">
  <div class="modal-content rounded-0">
    <div class="modal-header">
      <h5 class="modal-title fw-bold w-100 text-center" id="newsannouncementModalLabel">Edit Newsletter Announcement</h5>
      <!--<button type="button" class="btn-close rounded-0" data-bs-dismiss="modal" aria-label="Close"></button> -->
    </div>  
    <div class="modal-body announcement-modal"> 
      <div class="row align-items-center my-3"><div class="col-4">Title</div><div class="col-1">:</div><div class="col-7"><input class="form-control rounded-0" type="text" id="newsedittitle"></div></div>
      <!--<div class="row align-items-center my-3"><div class="col-4">Source</div><div class="col-1">:</div><div class="col-7 clsRadioSec">
        <label><input type="radio" class="Eradioc" name="EurlFile" id="newsEurlRadio" value="Url"> Url </label>
        <label><input type="radio"  class="Eradioc" name="EurlFile" id="newsEfileRadio" value="File"> File</label>
        </div></div>-->
      <div class="row align-items-center my-3" id="newsEurlSection" style="display:none"><div class="col-4">URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="newsediturl"></div></div>
      <div class="row align-items-start my-3" id="newsEfileSection" style="display:none"><div class="col-4">File</div><div class="col-1">:</div><div class="col-7" id="newseditFUploadSec"><div><input class="form-control-file custom-file-upload" type="file" id="newsuploadfileedit"></div><div class="uploadedFilenews mt-1"></div></div></div>
      <div class="row align-items-center my-3"><div class="col-4">Document/Url Properties</div><div class="col-1">:</div><div class="col-7">
  
      <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">

<input type="checkbox" class="btn-check" id="newseditsensitive" autocomplete="off">
<label class="btn btn-outline-theme" for="newseditsensitive">Sensitive</label>

<input type="checkbox" class="btn-check" id="newseditvisible" autocomplete="off">
<label class="btn btn-outline-theme" for="newseditvisible">Visible</label>

<input type="checkbox" class="btn-check" id="newseditnewtab" autocomplete="off">
<label class="btn btn-outline-theme" for="newseditnewtab">Open a new tab</label>
</div>

      </div></div>
    </div>
    <div class="modal-footer justify-content-between"> 
    <div>
    <button type="button" class="btn btn-sm btn-danger rounded-0" id="newsAnABtnDelete" data-bs-toggle="modal" data-bs-target="#newsAnADeleteModal">Delete</button>
     </div>
      <div class="d-flex">
      <button type="button" class="btn btn-sm btn-secondary mx-1 rounded-0"  id = "newsbtnUpdateClose" data-bs-dismiss="modal">Close</button>
      <button type="button" class="btn btn-sm btn-theme mx-1 rounded-0
      " id="newsbtnupdate">Update</button> </div>
    </div>
  </div>
</div>
</div> 

<div class="modal fade" id="newsAnADeleteModal" tabindex="-1" aria-labelledby="AnADeleteModalLabel" aria-hidden="true">
  <div class="modal-dialog AnA-delete-warning-dialog">
    <div class="modal-content rounded-0">
      <div class="modal-header">
        
      </div>
      <div class="modal-body AnA-delete-warning text-center pt-5"> 
      <h5 class="modal-title" id="newsdeleteAlterModalLabel">Confirmation</h5>
      <p class="mb-0">Are you sure want to Delete?</p>
      </div>
      <div class="modal-footer">
        <button type="button" id="newscancelAnADelete" class="btn btn-sm btn-secondary rounded-0" data-bs-dismiss="modal">No</button>
        <button type="button" id="newsconfirmAnADelete" class="btn btn-sm btn-danger rounded-0">Yes</button>
      </div>  
    </div>
  </div>
</div>

    <div class="border announcement-sec">
    <h5 class="bg-secondary text-light px-4 py-2" id="newsheaderTitle">Quarterly Newsletter</h5>
    <div class="add-announcements px-4 py-1 border-bottom"><a class="text-info cursor" data-bs-toggle="modal" data-bs-target="#newsannouncementModal">+ Add newsletter</a></div>
    <div id="newsannouncement-list" class="announcement-list">  
    <ul class="list-unstyled" id="newsannouncement-one"> 
    
    </ul> 
    </div>  
    </div>
    </div>
   
    </div> 
    </div>
    <!-- sensitive Modal -->
           
                <div class="modal fade" id="SensitiveModalnews" tabindex="-1" aria-labelledby="AnADeleteModalLabel" aria-hidden="true">
             <div class="modal-dialog sensitive-warning-dialog">
               <div class="modal-content rounded-0">
                 <div class="modal-header">
                    
                   <!-- <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>-->
                 </div>
                 <div class="modal-body sensitive-warning text-center pt-5"> 
                 <h5 class="modal-title" id="">Warning</h5>
                 <p class="mb-0">This is a sensitive document. Please don't share it externally.</p>
                 </div>
                 <div class="modal-footer">
                   <!--<button type="button" class="btn btn-sm btn-secondary rounded-0" data-bs-dismiss="modal">No</button>-->
                   <button type="button" id="btnyesnews" class="btn btn-sm btn-danger rounded-0" data-bs-dismiss="modal">OK</button>
                 </div>
               </div>
             </div>
           </div> 
           <!-- sensitive Modal -->
      `;
      //$("#newsheaderTitle").text(headerTitle)
    $("#newsAnABtnDelete").click(()=>{
      $(".announcement-modal-dialog").hide();
    });
    $("#newscancelAnADelete").click(()=>{
      $(".announcement-modal-dialog").show();
    });
    $("input[type=radio][name=urlFile]").change(function(e) {
      urlFile = e.currentTarget.value;
      console.log(urlFile);
      if(e.currentTarget.value == "Url"){
        $("#newsurlSection").show();
        $("#newsfileSection").hide();
      }else if (e.currentTarget.value == "File"){
        $("#newsurlSection").hide();
        $("#newsfileSection").show();
        $("#newstxturl").val("");
      }
          });
    $("input[type=radio][name=EurlFile]").change(function(e) {
            updateUrlFile = e.currentTarget.value;
            console.log(updateUrlFile);
            if(e.currentTarget.value == "Url"){
              $("#newsEurlSection").show();
              $("#newsEfileSection").hide();
            }else if (e.currentTarget.value == "File"){
              $("#newsEurlSection").hide();
              $("#newsEfileSection").show();
              $("#newsediturl").val("");
            }
                });

    getFLXNewsLetterAnnouncements();
    $(document).on('click','.sensitivenews',async function(e)
      {
        Sdata=$(this).attr("data-index"); 
        e.preventDefault();
      });
      $(document).on('click','#btnyesnews', function(e)
    {
    if(allitems[Sdata].Openanewtab==true)
    window.open(allitems[Sdata].Url, '_blank');
    else
    window.location.href = allitems[Sdata].Url;
    });

    $("#newsbtnsubmit").click(async function()
    {
    //   $(".announcement-modal-dialog").hide();
    // await addItems();
    if(urlFile == "Url"){
      if (mandatoryforaddItemsUrl()) {
        $(".announcement-modal-dialog").hide();
        await addItems();   
      } else {
        console.log("All fileds not filled");
      }
    }
    else{
      if (mandatoryforaddItems()) {
        $(".announcement-modal-dialog").hide();
        await addItems();   
      } else {
        console.log("All fileds not filled");
      }
    }
    
    });
    $("#newsbtnupdate").click(async function()
    {
    //   $(".announcement-modal-dialog").hide();
    // await updateItems();
    if(updateUrlFile=="Url"){
      if (mandatoryforupdateItemsUrl()) {
        $(".announcement-modal-dialog").hide();
        await updateItems();   
      } else {
        console.log("All fileds not filled");
      }
    }
    else{
      if (mandatoryforupdateItems()) {  
        $(".announcement-modal-dialog").hide();
        await updateItems();   
      } else {
        console.log("All fileds not filled");
      }
    }
    
    });
    $(document).on('click','.icon-edit-announce',async function()
    {
        // FileFormFolder
    editdata=$(this).attr("data-id"); 
    console.log(editdata); 
    
    $("#newsedittitle").val(allitems[editdata].Title);
    $("#newsediturl").val(allitems[editdata].Url); 
 
    if(allitems[editdata].UrlOrFile == "File"){
      SelectedFileName = allitems[editdata].Url.split('/').pop();
      $(".uploadedFilenews").html("");
      $(".uploadedFilenews").html(`<a href="${allitems[editdata].Url}">${SelectedFileName}</a>`);
    }else{
      $(".uploadedFilenews").html("")
    }
       
    if(allitems[editdata].UrlOrFile == "Url"){
      updateUrlFile = "Url"
      $("#newsEurlRadio").prop("checked",true);
      $("#newsEurlSection").show();
      $("#newsEfileSection").hide();
    } else{
      updateUrlFile = "File"
      $("#newsEfileRadio").prop("checked",true);
      $("#newsEurlSection").hide();
      $("#newsEfileSection").show();
      // $("newseditFUploadSec").html(`<input class="form-control-file custom-file-upload" type="file" id="newsuploadfileedit">`)
    }
    console.log(`InList : ${urlFile}`);
    
    $("#newseditsensitive").prop( "checked", allitems[editdata].SensitiveDocument);
    $("#newseditvisible").prop("checked",allitems[editdata].Visible);
    $("#newseditnewtab").prop("checked",allitems[editdata].Openanewtab);
    
    });
    
    $("#newsbtnclose").click(function()
    {
      $("#newstxttitle").val("");
      $("#newsbtnsensitive").val("");
      $("#newsbtnvisible").val("");  
      $("#newsbtnnewtab").val("");
      $("#newsuploadfile").val("");
      $("#newstxturl").val("");
      
      let radioReset = document.getElementsByName("urlFile");
      for(var i=0;i<radioReset.length;i++)
      radioReset[i]["checked"] = false;
    });
    $("#newsbtnUpdateClose").click(()=>{
      $("#newsuploadfileedit").val("")

    })
    $(document).on("change", "#newsuploadfile", function () {
      if ($(this)[0].files.length > 0) {
        for (let index = 0; index < $(this)[0].files.length; index++) {
          const file = $("#newsuploadfile")[0]["files"][index];
          Fileupload.push(file);
        }
        //$(this).val("");
        $(this).parent().find("label").text("Choose File");
      }
    });
    $(document).on("change", "#newsuploadfileedit", function () {
      
      if ($(this)[0].files.length > 0) {
        for (let index = 0; index < $(this)[0].files.length; index++) {
          const file = $("#newsuploadfileedit")[0]["files"][index];
          FileuploadEdit.push(file);
        }
        //$(this).val("");
        $(this).parent().find("label").text("Choose File");
        $(".uploadedFilenews").html("")
      }else{
        $(".uploadedFilenews").html(`<a href="${allitems[editdata].Url}">${SelectedFileName}</a>`);
      }
    });
    $("#newsconfirmAnADelete").click(()=>{
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


async function getFLXNewsLetterAnnouncements()
{
  $(".loader-section").show();

  // if(FLXNewsLetterAnnouncements)
  // {
    await sp.web.lists.getByTitle("FLXNewsLetterAnnouncements").items.select("*").filter("Visible eq '" + 1 + "'").get().then(async (item)=>
    {
  var htmlfornewsannouncement="";
  allitems=item;
  console.log(allitems);
  if(item.length  == 0){
    
    $("#newsannouncement-list").html(`<div class="text-center pt-5">No Items Available</div>`)
  }
  for(var i=0;i<item.length;i++){
    Filename.push(item[i].Url.split('/').pop());
    console.log("Filename");
  console.log(Filename);
  
    if(item[i].SensitiveDocument==true){
    if(item[i].Openanewtab==true){
      if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
  
  }
  
  else {
    if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivenews" data-bs-toggle="modal" data-bs-target="#SensitiveModalnews" data-index=${i}>${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
  }
  }
  
  else{
    if(item[i].Openanewtab==true){ 
      if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
  
  }
  
  else {
    if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfornewsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}</a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#newsannouncementModalEdit"></div></li>`;
      }
  } 
  }
   
  }
  $("#newsannouncement-one").html("");
  $("#newsannouncement-one").html(htmlfornewsannouncement);
    }).catch((error)=>
    {
      console.log(error);
    });
    $(".loader-section").hide();
  }
  // else{
  //   $("#announcement-one").html("");
  // $("#announcement-one").html(`<li class="py-2 px-4 d-flex align-items-center row">No data to display or Please select list name</li>`);
  // }
  // }

  

  async function addItems() {
    $(".loader-section").show();
    var requestdata = {}; 
     if (Fileupload.length > 0) {
      await Fileupload.map((filedata) => {
            sp.web
              .getFolderByServerRelativeUrl("/sites/FLXCommunity/AnnouncementDocument/FLXNewsLetterAnnouncements")
              .files.add(filedata.name, filedata, true)
              .then (function(data){
                console.log(data);
                requestdata = {
                  Title: $("#newstxttitle").val(),
                  // DocumentUrl:{
                  //   "__metadata": { type: "SP.FieldUrlValue" },
                  //   Description: "",
                  //   Url: $("#txturl").val(),
                
                  // },
                  Url:data.data.ServerRelativeUrl,
                  SensitiveDocument: $("#newsbtnsensitive").is(':checked') ? true : false,
                  Visible: $("#newsbtnvisible").is(':checked') ? true : false,
                  Openanewtab: $("#newsbtnnewtab").is(':checked') ? true : false,
                  UrlOrFile:urlFile
                };
                sp.web.lists
                .getByTitle("FLXNewsLetterAnnouncements")
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
        Title: $("#newstxttitle").val(),
        // DocumentUrl:{
        //   "__metadata": { type: "SP.FieldUrlValue" },
        //   Description: "",
        //   Url: $("#txturl").val(),
      
        // },
        Url:$("#newstxturl").val(),
        SensitiveDocument: $("#newsbtnsensitive").is(':checked') ? true : false,
        Visible: $("#newsbtnvisible").is(':checked') ? true : false,
        Openanewtab: $("#newsbtnnewtab").is(':checked') ? true : false,
        UrlOrFile:urlFile
      };
      sp.web.lists
      .getByTitle("FLXNewsLetterAnnouncements")
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
    $(".loader-section").hide();
  }

  async function updateItems() {
    $(".loader-section").show();
console.log(FileuploadEdit);

    var requestdata = {}; 
    var Id=allitems[editdata].ID;
     if (FileuploadEdit.length > 0) {
      await FileuploadEdit.map((filedata) => {
            sp.web
              .getFolderByServerRelativeUrl("/sites/FLXCommunity/AnnouncementDocument/FLXNewsLetterAnnouncements")
              .files.add(filedata.name, filedata, true)
              .then (function(data){
                console.log(data);
                requestdata = {
                  Title: $("#newsedittitle").val(),
                  // DocumentUrl:{
                  //   "__metadata": { type: "SP.FieldUrlValue" },
                  //   Description: "",
                  //   Url: $("#txturl").val(),
                  // },
                  Url:data.data.ServerRelativeUrl,
                  SensitiveDocument: $("#newseditsensitive").is(':checked') ? true : false,
                  Visible: $("#newseditvisible").is(':checked') ? true : false,
                  Openanewtab: $("#newseditnewtab").is(':checked') ? true : false,
                  UrlOrFile:updateUrlFile
                };
                sp.web.lists
                .getByTitle("FLXNewsLetterAnnouncements")
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
        Title: $("#newsedittitle").val(),
        // DocumentUrl:{
        //   "__metadata": { type: "SP.FieldUrlValue" },
        //   Description: "",
        //   Url: $("#txturl").val(),
        // },
        
        SensitiveDocument: $("#newseditsensitive").is(':checked') ? true : false,
        Visible: $("#newseditvisible").is(':checked') ? true : false,
        Openanewtab: $("#newseditnewtab").is(':checked') ? true : false,
        UrlOrFile:updateUrlFile
      };
      sp.web.lists
      .getByTitle("FLXNewsLetterAnnouncements")
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
      $(".uploadedFilenews").html(`<p class="text-danger">File Cannot be Empty</p>`)
    }
    else{
      requestdata = {
        Title: $("#newsedittitle").val(),
        Url:$("#newsediturl").val(),
        SensitiveDocument: $("#newseditsensitive").is(':checked') ? true : false,
        Visible: $("#newseditvisible").is(':checked') ? true : false,
        Openanewtab: $("#newseditnewtab").is(':checked') ? true : false,
        UrlOrFile:updateUrlFile
      };
      sp.web.lists
      .getByTitle("FLXNewsLetterAnnouncements")
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
    $(".loader-section").hide();
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
          $(".loader-section").hide();
          location.reload();
        },
      })
      
      .show()
      .setHeader("<div class='fw-bold alertifyConfirmation'>Confirmation</div> ")
      .set("closable", false);
  }
const deleteAnA = (id) =>{
   sp.web.lists.getByTitle("FLXNewsLetterAnnouncements").items.getById(id).delete().then(()=>{location.reload()}).catch((error)=>{alert("Error Occured");})
}
function mandatoryforaddItemsUrl() {
  var isAllvalueFilled = true;
  if (!$("#newstxttitle").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  } else if (!$("#newstxturl").val()) {
    alertify.error("Please Enter the url ");
    isAllvalueFilled = false;
  }
  // else if (!$("#newsuploadfile").val()) {
  //   alertify.error("Please upload file");
  //   isAllvalueFilled = false;  
  // }   
  return isAllvalueFilled;
}

function mandatoryforaddItems() {
  var isAllvalueFilled = true;
  if (!$("#newstxttitle").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  } 
  // else if (!$("#newstxturl").val()) {
  //   alertify.error("Please Enter the url ");
  //   isAllvalueFilled = false;
  // }
  else if (!$("#newsuploadfile").val()) {
    alertify.error("Please upload file");
    isAllvalueFilled = false;  
  }   
  return isAllvalueFilled;
}

function mandatoryforupdateItemsUrl() {
  var isAllvalueFilled = true;
  if (!$("#newsedittitle").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  } else if (!$("#newsediturl").val()) {
    alertify.error("Please Enter the url ");
    isAllvalueFilled = false;
  }
  // else if (!$("#newsuploadfileedit").val()) {
  //   alertify.error("Please upload file");
  //   isAllvalueFilled = false;  
  // }   
  return isAllvalueFilled;
}
function mandatoryforupdateItems() {
  var isAllvalueFilled = true;
  if (!$("#newsedittitle").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  }
  //  else if (!$("#newsediturl").val()) {
  //   alertify.error("Please Enter the url ");
  //   isAllvalueFilled = false;
  // }
  // else if (!$("#newsuploadfileedit").val()) {
  //   alertify.error("Please upload file");
  //   isAllvalueFilled = false;  
  // }   
  return isAllvalueFilled;
}
  


