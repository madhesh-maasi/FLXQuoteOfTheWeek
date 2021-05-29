import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './QuoteOfTheWeekWebPart.module.scss';
import * as strings from 'QuoteOfTheWeekWebPartStrings';
import { SPComponentLoader } from "@microsoft/sp-loader";
//declare var $;
//import "jquery";
SPComponentLoader.loadScript(
  // "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.4.min.js"
  "https://code.jquery.com/jquery-3.5.1.js"
);

import * as $ from "jquery";
import * as moment from 'moment';
//import "../../ExternalRef/Js/jquery.min.js";
import "../../ExternalRef/Js/bootstrap.js";
import { sp } from "@pnp/sp/presets/all";
import "../../ExternalRef/css/bootstrap.min.css";
import "../../ExternalRef/css/style.css"; 
import "../../ExternalRef/css/alertify.min.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");  


var siteURL = "";
var week1,week2,week3,week4,week5;
var IDarray=[];


export interface IQuoteOfTheWeekWebPartProps {  
  description: string;
}

export default class QuoteOfTheWeekWebPart extends BaseClientSideWebPart<IQuoteOfTheWeekWebPartProps> {
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
    <div class="quotes-section container container-sm container-lg contaoiner-md">

<div class="modal fade" id="quotesModal" tabindex="-1" aria-labelledby="quotesModalLabel" aria-hidden="true">
<div class="modal-dialog quote-modal">
<div class="modal-content rounded-0">
<div class="modal-header">
<h5 class="modal-title w-100 fw-bold text-center" id="quotesModalLabel">Add / Edit Quotes</h5>
<!--<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>-->
</div>
<div class="modal-body">  
<div class="row my-2 justify-content-center"><div class="col-1"></div><div class="col-4 text-center fw-bolder">Date</div><div class="col-7 text-center fw-bolder">Quote</div></div>
<div class="row align-items-start my-4 mx-2"><div class="col-1">1</div><div class="col-4"><input type="date" class="form-control disabledate rounded-0" id="add-date1" aria-describedby=""></div><div class="col-7"><textarea class="form-control rounded-0" id="add-quotes1" aria-describedby=""></textarea></div></div>
<div class="row align-items-start my-4 mx-2"><div class="col-1">2</div><div class="col-4"><input type="date" class="form-control disabledate rounded-0" id="add-date2" aria-describedby=""></div><div class="col-7"><textarea class="form-control rounded-0" id="add-quotes2" aria-describedby=""></textarea></div></div>
<div class="row align-items-start my-4 mx-2"><div class="col-1">3</div><div class="col-4"><input type="date" class="form-control disabledate rounded-0" id="add-date3" aria-describedby=""></div><div class="col-7"><textarea class="form-control rounded-0" id="add-quotes3" aria-describedby=""></textarea></div></div>
<div class="row align-items-start my-4 mx-2"><div class="col-1">4</div><div class="col-4"><input type="date" class="form-control disabledate rounded-0" id="add-date4" aria-describedby=""></div><div class="col-7"><textarea class="form-control rounded-0" id="add-quotes4" aria-describedby=""></textarea></div></div>
<div class="row align-items-start my-4 mx-2"><div class="col-1">5</div><div class="col-4"><input type="date" class="form-control disabledate rounded-0" id="add-date5" aria-describedby=""></div><div class="col-7"><textarea class="form-control rounded-0" id="add-quotes5" aria-describedby=""></textarea></div></div>

</div>   
<div class="modal-footer footer-edit-quoteofweek">
<button type="button" class="btn btn-sm btn-secondary rounded-0" data-bs-dismiss="modal" id="btnclose">Close</button>
<button type="button" class="btn btn-sm btn-theme rounded-0" id="btnsubmit">Submit</button>
</div>
</div>
</div>
</div>

<div class="modal fade" id="quoteseditModal" tabindex="-1" aria-labelledby="quoteseditModalLabel" aria-hidden="true">
<div class="modal-dialog quote-modal">
<div class="modal-content">
<h5 class="modal-title" id="quoteseditModalLabel">Edit Quotes</h5>
<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
</div>
<div class="modal-body">
<div class="row my-2 justify-content-center"><div class="col-1"></div><div class="col-4 text-center fw-bolder">Date</div><div class="col-7 text-center fw-bolder">Quote</div></div>
<div id="EditQuotesoftheweek"></div>
<!--<div class="row align-items-start my-2"><div class="col-1">1</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><textarea class="form-control" id="" aria-describedby=""></textarea></div></div>
<div class="row align-items-start my-2"><div class="col-1">2</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><textarea class="form-control" id="" aria-describedby=""></textarea></div></div>
<div class="row align-items-start my-2"><div class="col-1">3</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><textarea class="form-control" id="" aria-describedby=""></textarea></div></div>
<div class="row align-items-start my-2"><div class="col-1">4</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><textarea class="form-control" id="" aria-describedby=""></textarea></div></div>
<div class="row align-items-start my-2"><div class="col-1">5</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><textarea class="form-control" id="" aria-describedby=""></textarea></div></div>-->

</div>
<div class="modal-footer">
<button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Close</button>
<button type="button" class="btn btn-sm btn-theme" id="btnupdate">Update</button>
</div>
</div>
</div>
</div>

<div class="modal fade" id="quotesViewModal" tabindex="-1" aria-labelledby="quotesViewModalLabel" aria-hidden="true">
<div class="modal-dialog quote-modal">
<div class="modal-content rounded-0">
<div class="modal-header">
<h5 class="modal-title fw-bold w-100 text-center" id="quotesViewModalLabel">View Quotes</h5>
<!--<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button> -->
</div>
<div class="modal-body">
<div class="row my-2 justify-content-center">
<div class="col-1"></div>
<div class="col-4 text-center fw-bolder">Date</div><div class="col-7 text-center fw-bolder">Quote</div></div>
<div id="ViewQuotesoftheweek"></div>
<!--<div class="row align-items-start my-3"><div class="col-1">1</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><label>Sample</label></div></div>
<div class="row align-items-start my-3"><div class="col-1">2</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><label>Sample</label></div></div>
<div class="row align-items-start my-3"><div class="col-1">3</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><label>Sample</label></div></div>
<div class="row align-items-start my-3"><div class="col-1">4</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><label>Sample</label></div></div>
<div class="row align-items-start my-3"><div class="col-1">5</div><div class="col-4"><input type="date" class="form-control" id="" aria-describedby=""></div><div class="col-7"><label>Sample</label></div></div>-->
</div>
<div class="modal-footer">
<button type="button" class="btn btn-sm btn-secondary rounded-0" data-bs-dismiss="modal">Close</button>
<!--<button type="button" class="btn btn-sm btn-theme" data-bs-dismiss="modal" data-bs-toggle="modal" data-bs-target="#quoteseditModal">Edit</button>-->
</div>
</div>
</div>  
</div>

<div class="border bg-white">
<div class="tile-head bg-secondary p-2">
<h6 class="mx-2 mt-2" >Quote of the Week
</h6>
</div>
<div class="section-action my-2 mx-2 row align-items-center justify-content-end">
<div class="col p-0 d-flex align-items-center justify-content-end">
<span class="action-btn action-view mx-1" data-bs-toggle="modal" data-bs-target="#quotesViewModal"></span>
<span class="action-btn action-add mx-1" data-bs-toggle="modal" data-bs-target="#quotesModal" id="btnadd"></span>
</div>
</div>    
<div class="card p-4 m-2">
<div class="card-body p-4 ">
<span class="leftquote"></span>
<p class="card-text text-color text-center" id="quotes"></p>
<span class="rightquote text-center"></span>
</div>
</div>
</div>
</div>`;    
    getQuotesoftheWeek();  
    $("#btnsubmit").click(async function()
    {
      await addQuotes();
    });
    $("#btnupdate").click(async function()
    {
      await updateQuotes();
    });
    $("#btnclose").click(async function()
    {
      $("#add-quotes1").val("");
      $("#add-quotes2").val("");
      $("#add-quotes3").val("");
      $("#add-quotes4").val("");
      $("#add-quotes5").val("");
    });
    $("#btnadd").click(async function()
    {
      for(var i=0;i<IDarray.length;i++)
{  
if(IDarray[0]){
$("#add-quotes1").val(IDarray[0].Quotesoftheweek);
$("#add-quotes1").prop('data-intrusive','true')
  }
  if(IDarray[1]){
$("#add-quotes2").val(IDarray[1].Quotesoftheweek);
$("#add-quotes2").prop('data-intrusive','true')
  }
  if(IDarray[2]){
$("#add-quotes3").val(IDarray[2].Quotesoftheweek);
$("#add-quotes3").prop('data-intrusive','true')
  }
  if(IDarray[3]){
$("#add-quotes4").val(IDarray[3].Quotesoftheweek);
$("#add-quotes4").prop('data-intrusive','true')
  }
if(IDarray[4]){
$("#add-quotes5").val(IDarray[4].Quotesoftheweek);
$("#add-quotes5").prop('data-intrusive','true')
}

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


async function getQuotesoftheWeek()
{
  await sp.web.lists.getByTitle("Quotesoftheweek").items.select("*").get().then(async (item)=>
  {
    var today = new Date();
    console.log(item);
    for(var i=0;i<item.length;i++){
      var startdate=new Date(item[i].WeekStartDate);
      var sdate=new Date(item[i].WeekStartDate);
      var Edate=sdate.setDate(sdate.getDate() + 6);
var enddate=new Date(Edate);
      // if(today > startdate && today < enddate){
var startdatemt=moment(startdate).format("YYYY-MM-DD");
var enddatemt=moment(enddate).format("YYYY-MM-DD");
var todaymt=moment(today).format("YYYY-MM-DD");

      if(todaymt >= startdatemt && todaymt < enddatemt || todaymt > startdatemt && todaymt <= enddatemt){

$("#quotes").html("");  
$("#quotes").html(item[i].Quotesoftheweek); 
}
}

var d = new Date();
  var day = d.getDay(),
      diff = d.getDate() - day + (day == 0 ? -6:1); // adjust when day is sunday
  var currentdate=new Date(d.setDate(diff));
 week1=new Date(d.setDate(diff));
  console.log(currentdate);
$("#add-date1").val(moment(currentdate).format("YYYY-MM-DD"));
week2=currentdate. setDate(currentdate. getDate() + 7); 
$("#add-date2").val(moment(week2).format("YYYY-MM-DD"));
week3=currentdate. setDate(currentdate. getDate() + 7); 
$("#add-date3").val(moment(week3).format("YYYY-MM-DD"));
week4=currentdate. setDate(currentdate. getDate() + 7); 
$("#add-date4").val(moment(week4).format("YYYY-MM-DD"));
week5=currentdate. setDate(currentdate. getDate() + 7); 
$("#add-date5").val(moment(week5).format("YYYY-MM-DD"));
viewQuotes();    
  }).catch((error)=>
  {
    console.log(error);
  });
  }
  async function viewQuotes() {
    await sp.web.lists.getByTitle("Quotesoftheweek").items.select("*").get().then(async (item)=>
  {
    var htmlforviewquotes="";
    var htmlforeditquotes="";
    var count=0; 
    console.log(item);
    for(var i=0;i<item.length;i++){
      var startdate=new Date(item[i].WeekStartDate);
      var sdate=new Date(item[i].WeekStartDate);
      var Edate=sdate.setDate(sdate.getDate() + 6);
var enddate=new Date(Edate);  
var startdatemt=moment(startdate).format("YYYY-MM-DD");
var enddatemt=moment(enddate).format("YYYY-MM-DD");
var week1mt=moment(week1).format("YYYY-MM-DD");
var week2mt=moment(week2).format("YYYY-MM-DD");
var week3mt=moment(week3).format("YYYY-MM-DD");  
var week4mt=moment(week4).format("YYYY-MM-DD");
var week5mt=moment(week5).format("YYYY-MM-DD");

      if(week1mt >= startdatemt && week1mt < enddatemt || week2mt >= startdatemt && week2mt < enddatemt || week3mt >= startdatemt && week3mt < enddatemt || week4mt >= startdatemt && week4mt < enddatemt || week5mt >= startdatemt && week5mt < enddatemt){
        count++;
htmlforviewquotes+=`<div class="row align-items-start my-4 mx-2"><div class="col-1">${count}</div><div class="col-4"><input type="date" class="form-control disabledate rounded-0" id="" aria-describedby="" value="${startdatemt}"></div><div class="col-7 divlabel"><label>${item[i].Quotesoftheweek}</label></div></div>`;

htmlforeditquotes+=`<div class="row align-items-start my-2"><div class="col-1">${count}</div><div class="col-4"><input type="date" class="form-control disabledate" id="update-date${count}" aria-describedby="" value="${startdatemt}" ></div><div class="col-7"><textarea class="form-control update-quotes" id="update-quotes${count}" aria-describedby="" data-index=${count-1}>${item[i].Quotesoftheweek}</textarea></div></div>`;
IDarray.push({"ID":item[i].ID,"Quotesoftheweek":item[i].Quotesoftheweek});  
      }
    }    
 
$("#ViewQuotesoftheweek").html("");
$("#ViewQuotesoftheweek").html(htmlforviewquotes);

$("#EditQuotesoftheweek").html("");
$("#EditQuotesoftheweek").html(htmlforeditquotes);
for(var i=0;i<IDarray.length;i++)
{  
if(IDarray[0]){
$("#add-quotes1").val(IDarray[0].Quotesoftheweek);
$("#add-quotes1").prop('data-intrusive','true')
  }
  if(IDarray[1]){
$("#add-quotes2").val(IDarray[1].Quotesoftheweek);
$("#add-quotes2").prop('data-intrusive','true')
  }
  if(IDarray[2]){
$("#add-quotes3").val(IDarray[2].Quotesoftheweek);
$("#add-quotes3").prop('data-intrusive','true')
  }
  if(IDarray[3]){
$("#add-quotes4").val(IDarray[3].Quotesoftheweek);
$("#add-quotes4").prop('data-intrusive','true')
  }
if(IDarray[4]){
$("#add-quotes5").val(IDarray[4].Quotesoftheweek);
$("#add-quotes5").prop('data-intrusive','true')
}

}

disableallfields(); 
  }).catch((error) => {
    ErrorCallBack(error, "viewQuotes");
  });
}

  async function addQuotes() {

  let list = sp.web.lists.getByTitle('Quotesoftheweek');
  console.log(list);
  
    if($("#add-quotes1").prop('data-intrusive') && $("#add-quotes1").val()!=""){
    list.items.getById(IDarray[0].ID).update({ 
      WeekStartDate: $("#add-date1").val(),
      Quotesoftheweek:$("#add-quotes1").val()
    }).then(b => {
        console.log(b);
    });
  }
  else if($("#add-quotes1").val()!=""){
    list.items.add({ 
      WeekStartDate: $("#add-date1").val(),
      Quotesoftheweek:$("#add-quotes1").val()
    }).then(b => {
        console.log(b);
    });
  }
  if($("#add-quotes2").prop('data-intrusive') && $("#add-quotes2").val()!=""){
    list.items.getById(IDarray[1].ID).update({ 
      WeekStartDate: $("#add-date2").val(),
      Quotesoftheweek:$("#add-quotes2").val()
    }).then(b => {
        console.log(b);
    });
  } 
  else if($("#add-quotes2").val()!=""){
    list.items.add({ 
      WeekStartDate: $("#add-date2").val(),
      Quotesoftheweek:$("#add-quotes2").val()
    }).then(b => {
        console.log(b);
    });
  }
  if($("#add-quotes3").prop('data-intrusive') && $("#add-quotes3").val()!=""){
    list.items.getById(IDarray[2].ID).update({ 
      WeekStartDate: $("#add-date3").val(),
      Quotesoftheweek:$("#add-quotes3").val()
    }).then(b => {
        console.log(b);
    });
  }
  else if($("#add-quotes3").val()!=""){
    list.items.add({ 
      WeekStartDate: $("#add-date3").val(),
      Quotesoftheweek:$("#add-quotes3").val()
    }).then(b => {
        console.log(b);
    });
  }
  if($("#add-quotes4").prop('data-intrusive') && $("#add-quotes4").val()!=""){
    list.items.getById(IDarray[3].ID).update({ 
      WeekStartDate: $("#add-date4").val(),
      Quotesoftheweek:$("#add-quotes4").val()
    }).then(b => {
        console.log(b);
    });
  }
  else if($("#add-quotes4").val()!=""){
    list.items.add({ 
      WeekStartDate: $("#add-date4").val(),
      Quotesoftheweek:$("#add-quotes4").val()
    }).then(b => {
        console.log(b);
    });
  }
  if($("#add-quotes5").prop('data-intrusive') && $("#add-quotes5").val()!=""){
    list.items.getById(IDarray[4].ID).update({ 
      WeekStartDate: $("#add-date5").val(),
      Quotesoftheweek:$("#add-quotes5").val()
    }).then(b => {
        console.log(b);
    });
  }
  else if($("#add-quotes5").val()!=""){
    list.items.add({ 
      WeekStartDate: $("#add-date5").val(),
      Quotesoftheweek:$("#add-quotes5").val()
    }).then(b => {
        console.log(b);
    });
  }
  await  AlertMessage("<div class='alertfy-success'>Submitted successfully</div>");   
  }

  async function updateQuotes() {
    $('.update-quotes').each(function()
    {
    IDarray[$(this).attr('data-index')].Quotesoftheweek=$(this).val();
    });
    var count=1;
    var requesttaskdata = {};
                for(var i=0;i<IDarray.length;i++)
                {
                  var Id=IDarray[i].ID;
                  requesttaskdata = {
                    Quotesoftheweek: IDarray[i].Quotesoftheweek,
                    };
                    await sp.web.lists
                      .getByTitle("Quotesoftheweek")
                       .items.getById(Id)
                       .update(requesttaskdata).then(function (data) {             
                      count++;
                      if(count==IDarray.length)
                      {
                       AlertMessage("Something went wrong.please contact system admin"); 
                      }
                      
                    })
                    .catch(function (error) {
                      ErrorCallBack(error, "updateQuotes");
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
    .setHeader("<div class='fw-bold alertifyConfirmation'>Confirmation</div> ")
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

function disableallfields()
{
  $(".disabledate").prop('disabled',true);
}