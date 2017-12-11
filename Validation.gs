//此处定义全局变量
var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1U_kc73Mhzkl_J9XA91AQqRjClopQFmLj4E3jZdQKNr4/";
var SHEETS = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
var TITLE = SHEETS.getSheetByName("Configuration").getRange("B1").getValue();
var VALIDATION_URL = SHEETS.getSheetByName("Configuration").getRange("B4").getValue();
var FORM_ID = SHEETS.getSheetByName("Configuration").getRange("B5").getValue();


function doGet(e) {
  var parameters = e.parameter;
  var requestId = parameters.requestId;
  var responseId = parameters.responseId;
  var status = parameters.status;
  var myDate = new Date();
  var approver = parameters.approver;
  approver=approver.replace(/\s+/g,"");//去除空格，updated on 2017/11/10
  approver=approver.toLowerCase();//转换为小写，updated on 2017/11/17
  var decision = parameters.decision;
  var form = FormApp.openById(FORM_ID);
  var formResponse = form.getResponse(responseId);
  var itemResponses = formResponse.getItemResponses();
  var prefix = SHEETS.getSheetByName("Configuration").getRange("B2").getValue();
  var respondentEmail = formResponse.getRespondentEmail();
  var sheet;
  
  if (approver != Session.getActiveUser().getEmail()) return (HtmlService.createHtmlOutput("Sorry, you have no right to approve/reject."));
  
  sheet = SHEETS.getSheetByName("Validation History");
  
  var i;
  for (i = 2;; i++) {
    if (sheet.getRange("A" + i).getValue() == requestId && sheet.getRange("B" + i).getValue() == status) return (HtmlService.createHtmlOutput("You have already approved/rejected before."));
    if (sheet.getRange("A" + i).getValue() == "") break;
  }
  sheet.getRange("A" + i).setValue(requestId);
  sheet.getRange("B" + i).setValue(status);
  sheet.getRange("C" + i).setValue(approver);
  sheet.getRange("D" + i).setValue(decision);
  sheet.getRange("E" + i).setValue(myDate);
  var approverRole;
  if(status=="1"){
    approverRole=itemResponses[0].getItem().getTitle();
  }
  else approverRole=SHEETS.getSheetByName("Approval List").getRange(1,parseInt(status)).getValue();
  sheet.getRange("F" + i).setValue(approverRole);
  
  MailApp.sendEmail({
    to: respondentEmail,
    subject: TITLE + " - " + requestId,
    body: approver + " has reviewed your request. This is the step " + status + ".\nThe decision is: " + decision + ".", 
    noReply: true
  });
  
  if (decision == "Approve") {
    sheet = SHEETS.getSheetByName("Approval List");
    var nextApproverPosition = 1+parseInt(status);
    var nextApprover = sheet.getRange(2, nextApproverPosition, 1, 1).getValue();
    if (nextApprover) {
      var nextStatus = 1 + parseInt(status);
      
      var htmlBodyCode = "<!DOCTYPE HTML> <html>  <head>   <meta charset='utf-8'>   <meta name='Generator' content='Google Script'>   <meta name='Author' content='Yixiao Fei'>   <meta name='Keywords' content='web design, gather questions'>   <meta name='Description' content='Generate your own Quiz by filling the Spreadsheet'>   <title>IS Design | Welcome to Valeo</title>   <style>  table {    font-family: arial, sans-serif;    border-collapse: collapse;   width: 100%;}td, th {    border: 1px solid #dddddd;    text-align: left; padding: 8px;}tr:nth-child(even) {    background-color: #dddddd;}  body{     font: 10px/1.5 Arial, Helvetica, 'Microsoft YaHei', sans-serif;     padding: 0;     margin: 0;     background-color: #f4f4f4;     color: #818AA3;    }    .container{     width: 70%;     margin: auto;     overflow: hidden;    }    .logo{     margin-top: 15px;     float: left;    }    .title{     width: 80%;     margin: auto;     overflow: hidden;    }    .title h1{     float: right;     margin-top: 10px;     font-size: 20px;    }    .board{     background-color: #818AA3;     color: #ffffff;     text-align: right;     margin-bottom: 10px;    }    .button_1{     text-decoration: none;     border:0;     background-color: #2ECC71;     color: #ffffff;     padding: 4px 20px;     margin-top: 4px;    }    .button_1:hover{     color: #A2D9CE;     font-weight: bold;     }    footer p{     text-align: center;     color: #ffffff;     background-color: #818AA3;     padding: 20px;     margin-top: 20px;     font-size: 10px;                 font-style: normal;    }    section{     min-height: 500px;    }    .button_2{     text-decoration: none;     border:0;     background-color: #E74C3C;     color: #ffffff;     padding: 4px 20px;     margin-top: 4px;     width: 160px;    }    .button_2:hover{     color: #A2D9CE;     font-weight: bold;         }    form{     width: 70%;     margin: auto;     text-align: left;    }    form label{     font-size: 10px;     margin-top: 5px;     font-weight: bold;         }      #questionLabel{     float: left;    }    fieldset{     border-color: #2ECC71;    }   </style>  </head>  <body onload='fillInfo()'>   <header>   <div class='title'>    <div class='logo'><img style='height: 50px;' src='http://www.valeo.com/wp-content/themes/valeo/images/logo.png'></div>    ";
      htmlBodyCode += "<h1>"+TITLE+"</h1>";
      htmlBodyCode += "</div>   </header>   <section>    <div class='board'>     <h1 id='requestName'>";
      htmlBodyCode += ("Mail to requester: " + respondentEmail);
      htmlBodyCode += "</h1>    </div>    <div class='container'>     <fieldset>      <legend>Request Information</legend>      <form id='form1'>       ";
      
      //左右字段，用table保证对齐
      htmlBodyCode += "<br/><table>";
      for (var j = 0; j < itemResponses.length; j++) {
        var itemResponse = itemResponses[j];
        htmlBodyCode += "<tr>";
        htmlBodyCode += "<td>" + itemResponse.getItem().getTitle() + ": </td>";
        if(itemResponse.getItem().getType()=="FILE_UPLOAD"){
          var url = "https://drive.google.com/open?id=" + itemResponse.getResponse();
          htmlBodyCode += "<th><a href='" + url + "'>File Link</a></th>";
        }
        else htmlBodyCode += "<th>" + itemResponse.getResponse() + "</th>";
        htmlBodyCode += "</tr>";
      }
      htmlBodyCode += "</table><br/><br/>";     
      //btn1
      htmlBodyCode += "<div align='left' ><a class='button_1' id='approve' href='";
      htmlBodyCode += VALIDATION_URL+"?requestId=";
      htmlBodyCode += requestId;
      htmlBodyCode += "&responseId=";
      htmlBodyCode += responseId;
      htmlBodyCode += "&status=";
      htmlBodyCode += nextStatus;
      htmlBodyCode += "&approver=";
      htmlBodyCode += nextApprover;
      htmlBodyCode += "&decision=Approve";
      htmlBodyCode += "'>Approve</a>&nbsp;&nbsp;";
      //btn2
      htmlBodyCode += "<a class='button_1' id='approve_comment' href=\"";
      htmlBodyCode += "mailto:";
      htmlBodyCode += respondentEmail;
      htmlBodyCode += "?subject=外出%26车辆使用申请单 - " + requestId + " - Approve Comments";
      htmlBodyCode += "&body=";
      htmlBodyCode += "%0A%0AThank%20you%0ABest%20Regards";
      htmlBodyCode += "\">&nbsp;Approve Comment&nbsp;</a>&nbsp;&nbsp;<br/><br/>";
      //btn3
      htmlBodyCode += "<a class='button_2' id='reject' href='";
      htmlBodyCode += VALIDATION_URL+"?requestId=";
      htmlBodyCode += requestId;
      htmlBodyCode += "&responseId=";
      htmlBodyCode += responseId;
      htmlBodyCode += "&status=";
      htmlBodyCode += nextStatus;
      htmlBodyCode += "&approver=";
      htmlBodyCode += nextApprover;
      htmlBodyCode += "&decision=Reject";
      htmlBodyCode += "'>&nbsp;Reject&nbsp;</a>&nbsp;&nbsp;";
      //btn4
      htmlBodyCode += "<a class='button_2' id='reject_comment' href=\"";
      htmlBodyCode += "mailto:";
      htmlBodyCode += respondentEmail;
      htmlBodyCode += "?subject=外出%26车辆使用申请单 - " + requestId + " - Reject Comments";
      htmlBodyCode += "&body=";
      htmlBodyCode += "%0A%0AThank%20you%0ABest%20Regards";
      htmlBodyCode += "\">&nbsp;Reject Comment&nbsp;</a>";
      
      htmlBodyCode += "</div></section>	<footer><p>Designed by Valeo Niles IS Team, Copyright &copy; 2017</p></footer></body></html>";
      
      MailApp.sendEmail({
        to: nextApprover,
        subject: TITLE + " - " + requestId + " from: " + respondentEmail,
        htmlBody: htmlBodyCode, 
        noReply: true
      });
    } 
    else {
      var htmlBodyCode = "<!DOCTYPE HTML> <html>  <head>   <meta charset='utf-8'>   <meta name='Generator' content='Google Script'>   <meta name='Author' content='Yixiao Fei'>   <meta name='Keywords' content='web design, gather questions'>   <meta name='Description' content='Generate your own Quiz by filling the Spreadsheet'>   <title>IS Design | Welcome to Valeo</title>   <style>  table {    font-family: arial, sans-serif;    border-collapse: collapse;   width: 100%;}td, th {    border: 1px solid #dddddd;    text-align: left; padding: 8px;}tr:nth-child(even) {    background-color: #dddddd;}  body{     font: 10px/1.5 Arial, Helvetica, 'Microsoft YaHei', sans-serif;     padding: 0;     margin: 0;     background-color: #f4f4f4;     color: #818AA3;    }    .container{     width: 70%;     margin: auto;     overflow: hidden;    }    .logo{     margin-top: 15px;     float: left;    }    .title{     width: 80%;     margin: auto;     overflow: hidden;    }    .title h1{     float: right;     margin-top: 10px;     font-size: 20px;    }    .board{     background-color: #818AA3;     color: #ffffff;     text-align: right;     margin-bottom: 10px;    }    .button_1{     text-decoration: none;     border:0;     background-color: #2ECC71;     color: #ffffff;     padding: 4px 20px;     margin-top: 4px;    }    .button_1:hover{     color: #A2D9CE;     font-weight: bold;     }    footer p{     text-align: center;     color: #ffffff;     background-color: #818AA3;     padding: 20px;     margin-top: 20px;     font-size: 10px;                 font-style: normal;    }    section{     min-height: 500px;    }    .button_2{     text-decoration: none;     border:0;     background-color: #E74C3C;     color: #ffffff;     padding: 4px 20px;     margin-top: 4px;     width: 160px;    }    .button_2:hover{     color: #A2D9CE;     font-weight: bold;         }    form{     width: 70%;     margin: auto;     text-align: left;    }    form label{     font-size: 10px;     margin-top: 5px;     font-weight: bold;         }      #questionLabel{     float: left;    }    fieldset{     border-color: #2ECC71;    }   </style>  </head>  <body onload='fillInfo()'>   <header>   <div class='title'>    <div class='logo'><img style='height: 50px;' src='http://www.valeo.com/wp-content/themes/valeo/images/logo.png'></div>    ";
      htmlBodyCode += "<h1>"+TITLE+"</h1>";
      htmlBodyCode += "</div>   </header>   <section>    <div class='board'>     <h1>";
      htmlBodyCode += "Congratulations. Your request has been approved!";
      htmlBodyCode += "</h1>    </div>    <div class='container'>     <fieldset>      <legend>Request Information</legend>      <form id='form1'> <br/>";
      
      //左右字段，用table保证对齐
      htmlBodyCode += "<table>";
      for (var j = 0; j < itemResponses.length; j++) {
        var itemResponse = itemResponses[j];
        htmlBodyCode += "<tr>";
        htmlBodyCode += "<td>" + itemResponse.getItem().getTitle() + ": </td>";
        if(itemResponse.getItem().getType()=="FILE_UPLOAD"){
          var url = "https://drive.google.com/open?id=" + itemResponse.getResponse();
          htmlBodyCode += "<th><a href='" + url + "'>File Link</a></th>";
        }
        else htmlBodyCode += "<th>" + itemResponse.getResponse() + "</th>";
        htmlBodyCode += "</tr>";
      }
      htmlBodyCode += "</table><br/>";    
      
      //Validation History
      htmlBodyCode += "</table>";
      sheet=SHEETS.getSheetByName("Validation History");
      
      //打印标题
      htmlBodyCode += "<tr>";
      for(var i=1;;i++){
        if(sheet.getRange(1,i).getValue()=="") break;
        htmlBodyCode += "<th>" + sheet.getRange(1,i).getValue() + "</th>";
      }
      htmlBodyCode += "</tr>";
      
      //打印该申请单的审批记录
      for(var i=2;;i++){
        var id=sheet.getRange(i,1).getValue();
        if(id=="") break;
        if(id==requestId){
          htmlBodyCode += "<tr>";
          for(var j=1;;j++){
            if(sheet.getRange(i,j).getValue()=="") break;
            htmlBodyCode += "<td>" + sheet.getRange(i,j).getValue() + "</td>";
          }
          htmlBodyCode += "</tr>";
        }
      }
      
      
      htmlBodyCode += "</section>	<footer><p>Designed by Valeo Niles IS Team, Copyright &copy; 2017</p></footer></body></html>";
      
      MailApp.sendEmail({
        to: respondentEmail,
        subject: TITLE + " - " + requestId,
        htmlBody: htmlBodyCode, 
        noReply: true
      });
    }
    return HtmlService.createTemplateFromFile('Approve').evaluate();
  } else return HtmlService.createTemplateFromFile('Reject').evaluate();
}
