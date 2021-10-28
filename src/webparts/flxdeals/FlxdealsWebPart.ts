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
import * as moment from "moment";
import styles from './FlxdealsWebPart.module.scss';
import * as strings from 'FlxdealsWebPartStrings';
import { sp } from "@pnp/sp/presets/all";
import "../../ExternalRef/Css/style.css";
import "../../ExternalRef/Css/Bootstrap.min.css";
import "../../ExternalRef/js/Bootstrap.js";
import "../../ExternalRef/css/alertify.min.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");  

var siteURL = "";
var listUrl = "";
var Badgingdays="";
var docurl ="";
var Filename=[];
var Fileupload=[];
var FileuploadEdit=[];
var allitems=[];
var editdata='';
var urlFile = "";
var updateUrlFile = "";
var Sdata="";
var FilteredAdmin =[];
var currentuser = "";
export interface IFlxdealsWebPartProps {
  description: string;
}
let SelectedFileName = ""
export default class FlxdealsWebPart extends BaseClientSideWebPart<IFlxdealsWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });  
    }); 
  } 

  public render(): void {
    siteURL = this.context.pageContext.web.absoluteUrl;
    currentuser = this.context.pageContext.user.email;
    var siteindex = siteURL.toLocaleLowerCase().indexOf("sites");
    listUrl = siteURL.substr(siteindex - 1) + "/Lists/";
    this.domElement.innerHTML = `
    <div class="loader-section" style="display:none"> 
    <div class="loader"></div>  
    </div></div> 
    <div class="cont"> 
     
    <div class="row announcements-section">
    <div class="col-6 announcement p-0">

    <div class="modal fade" id="dealsannouncementModal" tabindex="-1" aria-labelledby="announcementModalLabel" aria-hidden="true">
  <div class="modal-dialog announcement-modal-dialog">
    <div class="modal-content rounded-0">
      <div class="modal-header">
        <h5 class="modal-title fw-bold w-100 text-center" id="dealsannouncementModalLabel">Add Document</h5>
      </div> 
      <div class="modal-body announcement-modal"> 
        <div class="row align-items-center my-3"><div class="col-4 titleannouncements">Title</div><div class="col-1">:</div><div class="col-7">
        <input class="form-control rounded-0" type="text" id="dealstxttitle"></div></div>
        
        
        <div class="row align-items-center my-3"><div class="col-4 titleannouncements">Source</div><div class="col-1">:</div><div class="col-7 clsRadioSec">
        
        <label><input type="radio" class="radioc" name="urlFile" id="dealsurlRadio" value="Url"> Url </label>
        <label><input type="radio"  class="radioc" name="urlFile" id="dealsfileRadio" value="File"> File</label>


        
        </div></div>


        <div class="row align-items-center my-3 radioToggle" id="dealsurlSection" style="display:none"><div class="col-4 titleannouncements">URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control rounded-0" type="text" id="dealstxturl"></div></div>
        <div class="row align-items-center my-3 radioToggle" id="dealsfileSection" style="display:none"><div class="col-4 titleannouncements">File</div><div class="col-1">:</div><div class="col-7"><input class="form-control-file custom-file-upload" type="file" id="dealsuploadfile"></div></div>
        <div class="row align-items-center my-3"><div class="col-4">Document/Url Properties</div><div class="col-1">:</div><div class="col-7">
  
        <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">

  <input type="checkbox" class="btn-check" id="dealsbtnsensitive" autocomplete="off">
  <label class="btn btn-outline-theme" for="dealsbtnsensitive">Sensitive</label>
  
  <input type="checkbox" class="btn-check" id="dealsbtnvisible" autocomplete="off">
  <label class="btn btn-outline-theme" for="dealsbtnvisible">Visible</label>

  <input type="checkbox" class="btn-check" id="dealsbtnnewtab" autocomplete="off">
  <label class="btn btn-outline-theme" for="dealsbtnnewtab">Open a new tab</label>
</div>
 
        </div></div>
      </div>
      <div class="modal-footer"> 
        <button type="button" class="btn btn-sm btn-secondary rounded-0" data-bs-dismiss="modal" id="dealsbtnclose">Close</button>
        <button type="button" class="btn btn-sm btn-theme  rounded-0" id="dealsbtnsubmit">Submit</button> 
      </div>
    </div>
  </div>
</div>   
  
<div class="modal fade" id="dealsannouncementModalEdit" tabindex="-1" aria-labelledby="announcementModalLabel" aria-hidden="true">
<div class="modal-dialog announcement-modal-dialog">
  <div class="modal-content rounded-0">
    <div class="modal-header">
      <h5 class="modal-title fw-bold w-100 text-center" id="dealsannouncementModalLabel">Edit Document</h5>
     <!-- <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button> -->
    </div>  
    <div class="modal-body announcement-modal"> 
      <div class="row align-items-center my-3"><div class="col-4 titleannouncements">Title</div><div class="col-1">:</div><div class="col-7"><input class="form-control rounded-0" type="text" id="dealsedittitle"></div></div>
      <!--<div class="row align-items-center my-3"><div class="col-4 titleannouncements">Source</div><div class="col-1">:</div><div class="col-7 clsRadioSec">
        <label><input type="radio" class="Eradioc" name="EurlFile" id="dealssEurlRadio" value="Url"> Url </label>
        <label><input type="radio"  class="Eradioc" name="EurlFile" id="dealsEfileRadio" value="File"> File</label>
        </div></div>-->
      <div class="row align-items-center my-3" id="dealsEurlSection" style="display:none"><div class="col-4 titleannouncements">URL</div><div class="col-1">:</div><div class="col-7"><input class="form-control" type="text" id="dealsediturl"></div></div>
      <div class="row align-items-start my-3" id="dealsEfileSection" style="display:none"><div class="col-4 titleannouncements">File</div><div class="col-1">:</div><div class="col-7" id="dealseditFUploadSec"><div><input class="form-control-file custom-file-upload" type="file" id="dealsuploadfileedit"></div><div class="uploadedFiledeals mt-1"></div></div></div>
      <div class="row align-items-center my-3"><div class="col-4">Document/Url Properties</div><div class="col-1">:</div><div class="col-7">
  
      <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">

<input type="checkbox" class="btn-check" id="dealseditsensitive" autocomplete="off">
<label class="btn btn-outline-theme" for="dealseditsensitive">Sensitive</label>

<input type="checkbox" class="btn-check" id="dealseditvisible" autocomplete="off">
<label class="btn btn-outline-theme" for="dealseditvisible">Visible</label>

<input type="checkbox" class="btn-check" id="dealseditnewtab" autocomplete="off">
<label class="btn btn-outline-theme" for="dealseditnewtab">Open a new tab</label>
</div>
   
      </div></div>
    </div>
    <div class="modal-footer justify-content-between"> 
    <div>
    <button type="button" class="btn btn-sm btn-danger  rounded-0" id="dealsAnABtnDelete" data-bs-toggle="modal" data-bs-target="#dealsAnADeleteModal">Delete</button>
     </div>
      <div class="d-flex">   
      <button type="button" class="btn btn-sm btn-secondary mx-1  rounded-0"  id="dealsbtnUpdateClose" data-bs-dismiss="modal">Close</button>
      <button type="button" class="btn btn-sm btn-theme mx-1 rounded-0" id="dealsbtnupdate">Update</button> </div>
    </div>
  </div>  
</div>
</div> 

<div class="modal fade" id="dealsAnADeleteModal" tabindex="-1" aria-labelledby="AnADeleteModalLabel" aria-hidden="true">
  <div class="modal-dialog AnA-delete-warning-dialog rounded-0">
    <div class="modal-content rounded-0">
      <div class="modal-header">
        
      </div>
      <div class="modal-body AnA-delete-warning text-center pt-5"> 
      <h5 class="modal-title" id="dealsdeleteAlterModalLabel">Confirmation</h5>
      <p class="mb-0">Are you sure want to Delete?</p>
      </div>
      <div class="modal-footer">
        <button type="button" id="dealscancelAnADelete" class="btn btn-sm btn-secondary  rounded-0" data-bs-dismiss="modal">No</button>
        <button type="button" id="dealsconfirmAnADelete" class="btn btn-sm btn-danger  rounded-0">Yes</button>
      </div>
    </div>
  </div>
</div>
<div class="viewallannounce d-flex justify-content-end">
<a href="#" class="info"  class="color-info"  id="ViewAlldeals">View All</a>
<a href="#" class="info"  class="color-info"  id="ShowVisibledeals">End User View</a>
    </div>
    <div class="border announcement-sec">
    <h5 class="bg-secondary text-light px-4 py-2" id="dealsheaderTitle">Sales, Placement & Proposal Reporting</h5>
    <div class="add-announcements px-4 py-1 border-bottom" id="add-deals"><a class="text-info cursor" data-bs-toggle="modal" data-bs-target="#dealsannouncementModal">+ Add document</a></div>
    <div id="dealsannouncement-list" class="announcement-list">  
    <ul class="list-unstyled" id="dealsannouncement-one"> 
    
    </ul> 
    </div>  
    </div>
    </div>  
   
    </div> 
    </div>
    <!-- sensitive Modal -->
           
                <div class="modal fade" id="SensitiveModaldeals" tabindex="-1" aria-labelledby="AnADeleteModalLabel" aria-hidden="true">
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
                   <button type="button" id="btnyesdeals" class="btn btn-sm btn-danger rounded-0" data-bs-dismiss="modal">OK</button>
                 </div>
               </div>
             </div>
           </div> 
           <!-- sensitive Modal -->  

           <div class="modal fade" id="exampleModalscroll2" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
           <div class="modal-dialog   modal-dialog-scrollable"">  
             <div class="modal-content rounded-0">
               <div class="modal-header">      
                 <h5 class="modal-title fw-bold w-100 text-center" id="exampleModalLabel">Sales, Placement & Proposal Reporting</h5>
             <!--   <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>  -->
               </div>     
               <div class="modal-body viewallmodal">
               <div class="viewallanounce">
               <ul class="list-unstyled">   
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>
               </li>     
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>  
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>  
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>  
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>  
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>  
               </li>
               <li class="d-flex align-items-center row my-3">
               <span class="announce-icon announce-icon announce-pdf col-1"></span><a href="#" class="col-8 p-0">FLX Announcements</a>
               <span class="icon-edit-announce col-2"></span>  
               </li>
               
              
               </ul> 
               </div>
               </div>
               <div class="modal-footer"> 
                 <button type="button" class="btn btn-sm btn-secondary rounded-0" data-bs-dismiss="modal" id="btnclose">Close</button>
                 <button type="button" class="btn btn-sm btn-theme rounded-0" id="btnsubmit">Submit</button> 
               </div>        
             </div>
           </div>
         </div>
          
    `;
    getadminfromsite();
    $("#ShowVisibledeals").hide();
      $("#ViewAlldeals").show();
      $("#ViewAlldeals").click(()=>{
        getFLXDealsAnnouncementsAll();
      });
      $("#ShowVisibledeals").click(()=>{
        getFLXDealsAnnouncements();
      });
    $("#dealsAnABtnDelete").click(()=>{
      $(".announcement-modal-dialog").hide();
    });
    $("#dealscancelAnADelete").click(()=>{
      $(".announcement-modal-dialog").show();
    });
    $("input[type=radio][name=urlFile]").change(function(e) {
      urlFile = e.currentTarget.value;
      console.log(urlFile);
      if(e.currentTarget.value == "Url"){
        $("#dealsurlSection").show();
        $("#dealsfileSection").hide();
      }else if (e.currentTarget.value == "File"){
        $("#dealsurlSection").hide();
        $("#dealsfileSection").show();
        $("#dealstxturl").val("");
      }
          });
    $("input[type=radio][name=EurlFile]").change(function(e) {
            updateUrlFile = e.currentTarget.value;
            console.log(updateUrlFile);
            if(e.currentTarget.value == "Url"){
              $("#dealsEurlSection").show();
              $("#dealsEfileSection").hide();
            }else if (e.currentTarget.value == "File"){
              $("#dealsEurlSection").hide();
              $("#dealsEfileSection").show();
              $("#dealsediturl").val("");
            }
                });

    $(document).on('click','.sensitivedeals',async function(e)
      {
        Sdata=$(this).attr("data-index"); 
        e.preventDefault();
      });
      $(document).on('click','#btnyesdeals', function(e)
    {
    if(allitems[Sdata].Openanewtab==true)
    window.open(allitems[Sdata].Url, '_blank');
    else
    window.location.href = allitems[Sdata].Url;
    });

    $("#dealsbtnsubmit").click(async function()
    { 
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
      // if (mandatoryforaddItems()) {
      //   $(".announcement-modal-dialog").hide();
      //   await addItems();   
      // } else {
      //   console.log("All fileds not filled");
      // }
      
    // await addItems();
    });   
    $("#dealsbtnupdate").click(async function()
    {
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
      
    //   $(".announcement-modal-dialog").hide();
    // await updateItems();
    });
    $(document).on('click','.icon-edit-announce',async function()
    {
        // FileFormFolder
    editdata=$(this).attr("data-id"); 
    console.log(editdata); 
    
    $("#dealsedittitle").val(allitems[editdata].Title);
    $("#dealsediturl").val(allitems[editdata].Url); 
 
    if(allitems[editdata].UrlOrFile == "File"){
      SelectedFileName = allitems[editdata].Url.split('/').pop();
      $(".uploadedFiledeals").html("");
      $(".uploadedFiledeals").html(`<a href="${allitems[editdata].Url}">${SelectedFileName}</a>`);
    }else{
      $(".uploadedFiledeals").html("")
    }
       
    if(allitems[editdata].UrlOrFile == "Url"){
      updateUrlFile = "Url"
      $("#dealsEurlRadio").prop("checked",true);
      $("#dealsEurlSection").show();
      $("#dealsEfileSection").hide();
    } else{
      updateUrlFile = "File"
      $("#dealsEfileRadio").prop("checked",true);
      $("#dealsEurlSection").hide();
      $("#dealsEfileSection").show();
      // $("dealseditFUploadSec").html(`<input class="form-control-file custom-file-upload" type="file" id="dealsuploadfileedit">`)
    }
    console.log(`InList : ${urlFile}`);
    
    $("#dealseditsensitive").prop( "checked", allitems[editdata].SensitiveDocument);
    $("#dealseditvisible").prop("checked",allitems[editdata].Visible);
    $("#dealseditnewtab").prop("checked",allitems[editdata].Openanewtab);
    
    });
    
    $("#dealsbtnclose").click(function()
    {
      $("#dealstxttitle").val("");
      $("#dealsbtnsensitive").val("");
      $("#dealsbtnvisible").val("");
      $("#dealsbtnnewtab").val("");
      $("#dealsuploadfile").val("");
      $("#dealstxturl").val("");
      
      let radioReset = document.getElementsByName("urlFile");
      for(var i=0;i<radioReset.length;i++)
      radioReset[i]["checked"] = false;
    });
    $("#dealsbtnUpdateClose").click(()=>{
      $("#dealsuploadfileedit").val("")

    })
    $(document).on("change", "#dealsuploadfile", function () {
      if ($(this)[0].files.length > 0) {
        for (let index = 0; index < $(this)[0].files.length; index++) {
          const file = $("#dealsuploadfile")[0]["files"][index];
          Fileupload.push(file);
        }
        //$(this).val("");
        $(this).parent().find("label").text("Choose File");
      }
    });
    $(document).on("change", "#dealsuploadfileedit", function () {
      
      if ($(this)[0].files.length > 0) {
        for (let index = 0; index < $(this)[0].files.length; index++) {
          const file = $("#dealsuploadfileedit")[0]["files"][index];
          FileuploadEdit.push(file);
        }
        //$(this).val("");
        $(this).parent().find("label").text("Choose File");
        $(".uploadedFiledeals").html("")
      }else{
        $(".uploadedFiledeals").html(`<a href="${allitems[editdata].Url}">${SelectedFileName}</a>`);
      }
    });
    $("#dealsconfirmAnADelete").click(()=>{
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
async function getadminfromsite() {

  var bag=[];
  let listLocation  = await sp.web.getList(listUrl + "Badging").items.get(); 
  listLocation.forEach((li) => {
   bag.push(li.Days); 
   console.log(bag);
  });
  Badgingdays= bag[0];
  console.log(Badgingdays);

  var AdminInfo = [];
  await sp.web.siteGroups
    .getByName("FLX Admins")
    .users.get()
    .then(function (result) {
      for (var i = 0; i < result.length; i++) {
        AdminInfo.push({
          Title: result[i].Title,
          ID: result[i].Id,
          Email: result[i].Email,
        });
      }
      FilteredAdmin = AdminInfo.filter((admin)=>{return (admin.Email == currentuser)});
      console.log(FilteredAdmin);
      getFLXDealsAnnouncements();
    })
    .catch(function (err) {
      console.log(err);
      //alert("Group not found: " + err);
    });
}
async function getFLXDealsAnnouncements()
{
  $("#ShowVisibledeals").hide();
  $("#ViewAlldeals").show();
  allitems=[];
  Filename=[];

  // if(FLXDealsAnnouncements)
  // {
    await sp.web.lists.getByTitle("FLXDealsAnnouncements").items.select("*").filter("Visible eq '" + 1 + "'").orderBy("Modified",false).get().then(async (item)=>
    {
  var htmlfordealsannouncement="";
  allitems=item;
  console.log(allitems);
  if(item.length  == 0){
    
    $("#dealsannouncement-list").html(`<div class="text-center pt-5">No Items Available</div>`)
  }
  for(var i=0;i<item.length;i++){
    Filename.push(item[i].Url.split('/').pop());   
    console.log("Filename");
  console.log(Filename);
  
    if(item[i].SensitiveDocument==true){
    if(item[i].Openanewtab==true){
      if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt  mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
  
  }
  
  else {
    if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#sdealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
  }
  }
  
  else{
    if(item[i].Openanewtab==true){ 
      if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
  
  }
  
  else {
    if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsnnouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
  } 
  }
   
  }
  if (FilteredAdmin.length>0) 
        {
          $("#dealsannouncement-one").html("");
          $("#dealsannouncement-one").html(htmlfordealsannouncement);
        }
        else{
          $("#dealsannouncement-one").html("");
          $("#dealsannouncement-one").html(htmlfordealsannouncement);
          $("#ViewAlldeals").hide();
          $("#ShowVisibledeals").hide();
          $(".icon-edit-announce").hide();
          $("#add-deals").hide();  
        }

        var count;
        for(var i=0;i<item.length;i++){
          count=i;
          var today = new Date();
          var startdate=new Date(item[i].Created);
              var sdate=new Date(item[i].Created);
              var Edate=sdate.setDate(sdate.getDate() + parseInt(Badgingdays));
        var enddate=new Date(Edate);
        var startdatemt=moment(startdate).format("YYYY-MM-DD");
        var enddatemt=moment(enddate).format("YYYY-MM-DD");
        var todaymt=moment(today).format("YYYY-MM-DD");
        
              if(todaymt >= startdatemt && todaymt < enddatemt || todaymt > startdatemt && todaymt <= enddatemt){
        
        $(".newvisdeals"+count).show();   
        }
        else{
          $(".newvisdeals"+count).hide(); 
        }
        }

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
    $(".loader-section").show();

    var requestdata = {}; 
     if (Fileupload.length > 0) {
      await Fileupload.map((filedata) => {
            sp.web
              .getFolderByServerRelativeUrl("/sites/FLXCommunity/AnnouncementDocument/FLXDealsAnnouncements")
              .files.add(filedata.name, filedata, true)
              .then (function(data){
                console.log(data);
                requestdata = {
                  Title: $("#dealstxttitle").val(),
                  // DocumentUrl:{
                  //   "__metadata": { type: "SP.FieldUrlValue" },
                  //   Description: "",
                  //   Url: $("#txturl").val(),
                
                  // },
                  Url:data.data.ServerRelativeUrl,
                  SensitiveDocument: $("#dealsbtnsensitive").is(':checked') ? true : false,
                  Visible: $("#dealsbtnvisible").is(':checked') ? true : false,
                  Openanewtab: $("#dealsbtnnewtab").is(':checked') ? true : false,
                  UrlOrFile:urlFile
                };
                sp.web.lists
                .getByTitle("FLXDealsAnnouncements")
                .items.add(requestdata)
                .then(function (data) {
                  console.log(data);
                  AlertMessage("<div class='alertfy-success'>Submitted successfully</div>");
                })
                .catch(function (error) {
                  // ErrorCallBack(error, "addItems");
                  console.log(error);
                  $(".loader-section").hide();
                });
              })
          });
    }  
    else{
      requestdata = {
        Title: $("#dealstxttitle").val(),
        // DocumentUrl:{
        //   "__metadata": { type: "SP.FieldUrlValue" },
        //   Description: "",
        //   Url: $("#txturl").val(),
      
        // },
        Url:$("#dealstxturl").val(),
        SensitiveDocument: $("#dealsbtnsensitive").is(':checked') ? true : false,
        Visible: $("#dealsbtnvisible").is(':checked') ? true : false,
        Openanewtab: $("#dealsbtnnewtab").is(':checked') ? true : false,
        UrlOrFile:urlFile
      };
      sp.web.lists
      .getByTitle("FLXDealsAnnouncements")
      .items.add(requestdata)
      .then(function (data) {
        console.log(data);
        AlertMessage("<div class='alertfy-success'>Submitted successfully</div>");
      })
      .catch(function (error) {
        // ErrorCallBack(error, "addItems");
        console.log(error);
        $(".loader-section").hide();
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
              .getFolderByServerRelativeUrl("/sites/FLXCommunity/AnnouncementDocument/FLXDealsAnnouncements")
              .files.add(filedata.name, filedata, true)
              .then (function(data){
                console.log(data);
                requestdata = {
                  Title: $("#dealsedittitle").val(),
                  // DocumentUrl:{
                  //   "__metadata": { type: "SP.FieldUrlValue" },
                  //   Description: "",
                  //   Url: $("#txturl").val(),
                  // },
                  Url:data.data.ServerRelativeUrl,
                  SensitiveDocument: $("#dealseditsensitive").is(':checked') ? true : false,
                  Visible: $("#dealseditvisible").is(':checked') ? true : false,
                  Openanewtab: $("#dealseditnewtab").is(':checked') ? true : false,
                  UrlOrFile:updateUrlFile
                };
                sp.web.lists
                .getByTitle("FLXDealsAnnouncements")
                .items.getById(Id).update(requestdata)
                .then(function (data) {
                  console.log(data);
                  AlertMessage("<div class='alertfy-success'>Updated successfully</div>");
                })
                .catch(function (error) {
                  // ErrorCallBack(error, "updateItems");
                  console.log(error);
                  $(".loader-section").hide();
                });
              })
          }); 
    } else if(FileuploadEdit.length == 0 && updateUrlFile == "File" && SelectedFileName != ""){
      requestdata = {
        Title: $("#dealsedittitle").val(),
        // DocumentUrl:{
        //   "__metadata": { type: "SP.FieldUrlValue" },
        //   Description: "",
        //   Url: $("#txturl").val(),
        // },
        
        SensitiveDocument: $("#dealseditsensitive").is(':checked') ? true : false,
        Visible: $("#dealseditvisible").is(':checked') ? true : false,
        Openanewtab: $("#dealseditnewtab").is(':checked') ? true : false,
        UrlOrFile:updateUrlFile
      };
      sp.web.lists
      .getByTitle("FLXDealsAnnouncements")
      .items.getById(Id).update(requestdata)
      .then(function (data) {
        console.log(data);
        AlertMessage("<div class='alertfy-success'>Updated successfully</div>");
      })
      .catch(function (error) {
        // ErrorCallBack(error, "updateItems");
        console.log(error);
        $(".loader-section").hide();
      });
    }else if(FileuploadEdit.length == 0 && updateUrlFile == "File" && SelectedFileName == ""){
      $(".uploadedFiledeals").html(`<p class="text-danger">File Cannot be Empty</p>`)
    }
    else{
      requestdata = {
        Title: $("#dealsedittitle").val(),
        Url:$("#dealsediturl").val(),
        SensitiveDocument: $("#dealseditsensitive").is(':checked') ? true : false,
        Visible: $("#dealseditvisible").is(':checked') ? true : false,
        Openanewtab: $("#dealseditnewtab").is(':checked') ? true : false,
        UrlOrFile:updateUrlFile
      };
      sp.web.lists
      .getByTitle("FLXDealsAnnouncements")
      .items.getById(Id).update(requestdata)
      .then(function (data) {
        console.log(data);
        AlertMessage("<div class='alertfy-success'>Updated successfully</div>");
      })
      .catch(function (error) {
        // ErrorCallBack(error, "updateItems");
        console.log(error);
        $(".loader-section").hide();
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
   sp.web.lists.getByTitle("FLXDealsAnnouncements").items.getById(id).delete().then(()=>{location.reload()}).catch((error)=>{alert("Error Occured");})
}
function mandatoryforaddItemsUrl() {
  var isAllvalueFilled = true;
  if (!$("#dealstxttitle").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  } else if (!$("#dealstxturl").val()) {
    alertify.error("Please Enter the url ");
    isAllvalueFilled = false;
  }
  // else if (!$("#dealsuploadfile").val()) {
  //   alertify.error("Please upload file");
  //   isAllvalueFilled = false;  
  // }   
  return isAllvalueFilled;
}
function mandatoryforaddItems() {
  var isAllvalueFilled = true;
  if (!$("#dealstxttitle").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  } 
  // else if (!$("#dealstxturl").val()) {
  //   alertify.error("Please Enter the url ");
  //   isAllvalueFilled = false;
  // }
  else if (!$("#dealsuploadfile").val()) {
    alertify.error("Please upload file");
    isAllvalueFilled = false;  
  }   
  return isAllvalueFilled;
}

function mandatoryforupdateItemsUrl() {
  var isAllvalueFilled = true;
  if (!$("#dealsedittitle").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  } else if (!$("#dealsediturl").val()) {
    alertify.error("Please Enter the url ");
    isAllvalueFilled = false;
  }
  // else if (!$("#dealsuploadfileedit").val()) {
  //   alertify.error("Please upload file");
  //   isAllvalueFilled = false;  
  // }   
  return isAllvalueFilled;
}
function mandatoryforupdateItems() {
  var isAllvalueFilled = true;
  if (!$("#dealsedittitle").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  } 
  // else if (!$("#dealsediturl").val()) {
  //   alertify.error("Please Enter the url ");
  //   isAllvalueFilled = false;
  // }
  // else if (!$("#dealsuploadfileedit").val()) {
  //   alertify.error("Please upload file");
  //   isAllvalueFilled = false;  
  // }   
  return isAllvalueFilled;
}

async function getFLXDealsAnnouncementsAll()
{
  $("#ShowVisibledeals").show();
  $("#ViewAlldeals").hide();
  allitems=[];
  Filename=[];

  // if(FLXDealsAnnouncements)
  // {
    await sp.web.lists.getByTitle("FLXDealsAnnouncements").items.select("*").orderBy("Modified",false).get().then(async (item)=>
    {
  var htmlfordealsannouncement="";
  allitems=item;
  console.log(allitems);
  if(item.length  == 0){
    
    $("#dealsannouncement-list").html(`<div class="text-center pt-5">No Items Available</div>`)
  }
  for(var i=0;i<item.length;i++){
    Filename.push(item[i].Url.split('/').pop());   
    console.log("Filename");
  console.log(Filename);
  
    if(item[i].SensitiveDocument==true){
    if(item[i].Openanewtab==true){
      if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt  mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
  
  }
  
  else {
    if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#sdealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" class="col-8 sensitivedeals" data-bs-toggle="modal" data-bs-target="#SensitiveModaldeals" data-index=${i}>${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
  }
  }
  
  else{
    if(item[i].Openanewtab==true){ 
      if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a data-interception="off" href="${item[i].Url}" target="_blank" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
  
  }
  
  else {
    if (Filename[i].split(".").pop() == "pdf")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-pdf mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsnnouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "ppt" || Filename[i].split(".").pop() == "pptx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-ppt mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "doc" || Filename[i].split(".").pop() == "docx")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-doc mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "xlsx" || Filename[i].split(".").pop() == "csv")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-excel mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else if (Filename[i].split(".").pop() == "png" || Filename[i].split(".").pop() == "jpg" || Filename[i].split(".").pop() == "jpeg")
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-img mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
      else
      {
        htmlfordealsannouncement+=`<li class="py-2 px-4 d-flex align-items-center row"><span class="announce-icon announce-new mx-1 col-1"></span><a href="${item[i].Url}" class="col-8">${item[i].Title}<span class="spannew newvisdeals${i}">New</span></a><div class="icon-edit-announce col-2" data-id=${i} data-bs-toggle="modal" data-bs-target="#dealsannouncementModalEdit"></div></li>`;
      }
  } 
  }
   
  }
  $("#dealsannouncement-one").html("");
  $("#dealsannouncement-one").html(htmlfordealsannouncement);

  var count;
  for(var i=0;i<item.length;i++){
    count=i;
    var today = new Date();
    var startdate=new Date(item[i].Created);
        var sdate=new Date(item[i].Created);
        var Edate=sdate.setDate(sdate.getDate() + parseInt(Badgingdays));
  var enddate=new Date(Edate);
  var startdatemt=moment(startdate).format("YYYY-MM-DD");
  var enddatemt=moment(enddate).format("YYYY-MM-DD");
  var todaymt=moment(today).format("YYYY-MM-DD");
  
        if(todaymt >= startdatemt && todaymt < enddatemt || todaymt > startdatemt && todaymt <= enddatemt){
  
  $(".newvisdeals"+count).show();   
  }
  else{
    $(".newvisdeals"+count).hide(); 
  }
  }

    }).catch((error)=>
    {
      console.log(error);
    });

  }