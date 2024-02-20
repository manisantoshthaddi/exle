/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// document.getElementById('mani').addEventListener('click',openDialog)
// function openDialog() {    
//   // var inputElement = document.getElementById("userInput");
//   // var code = inputElement.value;  
//   const axios = require('axios');
// const qs = require('qs');
// let data = qs.stringify({
//   'code': 'mP45874HWhkdKb7Yc8jPCg9hD',
//   'pu': 'https://mingle-sso.inforcloudsuite.com:443/DEVMRKT_DEV/as/',
//   'ot': 'token.oauth2',
//   'ci': 'DEVMRKT_DEV~akVifuI2mKwJxoe1qNpOgSkr7c8dLcjVm9WsHBjm-s8',
//   'cs': 'knco9yd-pBd_qaBnAiZwrVX3jnyMe4ehAb2U3kWrDY5QnVJ7hr7xE6DpWXPw2xmkQJkRK9WmmAhkqk8rZi2G2w',
//   'ru': 'https://localhost:3000/commands.html' 
// });

// let config = {
//   method: 'post',
//   maxBodyLength: Infinity,
//   url: 'https://devserveraddin.azurewebsites.net/commands/accesstoken',
//   headers: { 
//     'Content-Type': 'application/x-www-form-urlencoded'
//   },
//   data : data
// };

// axios.request(config)
// .then((response) => {
//   console.log(JSON.stringify(response.data));
// })
// .catch((error) => {
//   console.log(error);
// });

// }
import { off } from 'process';

/* global console, document, Excel, Office */

/* 
    Authors: Raghavender Hariharan, Singuri Suchith, Rohit Bhrugumalla, Ujwala Parupudi
    Team: Platform Technology Group
    Description: This code is used to develop an Excel Add-In that lets the user upload their 
    worksheets into ION via IMS (ION V2 Messaging Service) or into Infor Datalake via Data fabric 
    ingestion APIs. Users can also retrieve data from Infor Datalake into Excel Worksheets.
    Taskpane Folder: Contains the HTML and Javascript code that handles all the UI and functionalities
    for the Excel Add-In. 
*/

var _dlg;              // var for Dialog box
var error_var = 0;    // var to check if Log Sheet exists
var access_token="eyJraWQiOiJrZzplMGRlNmFiZC1jY2NlLTQyYzEtYjFlNS05ZDkwNjdhMmRkMGMiLCJhbGciOiJSUzI1NiJ9.eyJTZXJ2aWNlQWNjb3VudCI6IkRFVk1SS1RfREVWI01IclRocVhudlJGVWtMdVFwbU9zd0J0VG9fclhWUGNyc2t3ZEZvQTJBN1g0NnJ3LUladVRPMVZMNVJEa2I5aWE5ZTh5YzEycFU4RzRyUi1RbnlBNzVBIiwiVGVuYW50IjoiREVWTVJLVF9ERVYiLCJJZGVudGl0eTIiOiI0OTZlMGYxNi1lNjllLTQ1NmQtODgwYS1jZmFmZjkyMmNkMDEiLCJFbmZvcmNlU2NvcGVzRm9yQ2xpZW50IjoiMCIsImdyYW50X2lkIjoiYjZhMDk5NjAtODdjMS00NDU5LTliZjUtNzg0ZGQzMDVjOTBiIiwiSW5mb3JTVFNJc3N1ZWRUeXBlIjoiQVMiLCJjbGllbnRfaWQiOiJERVZNUktUX0RFVn5MXzNrdWs2M3ZlZTJHS2oyaHVXMmd1cFdKWU4zODg4TUNGZUtDWDFBN1ZrIiwianRpIjoiYjE0Y2E4ZTItZmU1NS00YWNiLTgxMTctYTVkMjQzMWViOTEyIiwiaWF0IjoxNzA1MzgyNTYyLCJuYmYiOjE3MDUzODI1NjIsImV4cCI6MTcwNTM4OTc2MiwiaXNzIjoiaHR0cHM6Ly9taW5nbGUtc3NvLmluZm9yY2xvdWRzdWl0ZS5jb206NDQzIiwiYXVkIjoiaHR0cHM6Ly9taW5nbGUtaW9uYXBpLmluZm9yY2xvdWRzdWl0ZS5jb20ifQ.YhFSwrI7iVPydSl5jnD9zaK8AlupO-Xfd5AhtctmxyfrOzStlfPT3oF1pp7PpglZlZeyDuY8CtioMXgkBTThPfoxf7wlxltP3qsEAgiOUd2vv2JQOMdimv2I-biDq-TZwT7DTz5D5ikjIhHMn_SFYF2tzgyO7CCN_-4R34RcuT04YKLy-9bCcPslSeSCiSubPZWYjxsR_w4q45bNZ_ZXjIPxGBSjJbRWlUSlRQIPmrKAqwxauJORDr2Wa9rDjzTiY9opKFG7fKNwofTeQyWa5Z2CJJI_nojcqi6NmGotHNAXeos6HKp0v1XVt9YUAnsNmu0ItSKHah4dstmAxAifug";     // var for storing Bearer token
var no_of_rows;       // var for storing rows in Log Sheet
var tenant="DEVMRKT_DEV";           // var for storing tenant value
var endpoint_url="https://mingle-ionapi.inforcloudsuite.com";     // var for the endpoint url
var logout_url = "";  // var for log out URL
var lid;              // var for storing logical id value
var color;            // var for color used in log sheet
var querynum = 0;          //var to represent the query number in Add Favourites

// Load Fetch Library
const fetch = (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));
const Papa = require('papaparse');

// Window OnLoad
window.addEventListener('load', addSheet);

// Add Status_Overview Sheet
export async function addSheet() {
  try {
    await Excel.run(async (context) => {
      // Add log sheet
      let sheets = context.workbook.worksheets;
      let log_sheet = sheets.add("Sheet_Overview");
      log_sheet.load("name, position");
      await context.sync();

      // Add Log Sheet Headers
      log_sheet = sheets.getItem("Sheet_Overview");
      let headers = [
        ["Sheet_Name", "Object_Schema", "Size(in Bytes)", "No of Rows", "Date", "Time", "                                 Status                                 ", "Error-Message"],
      ];
      let range = log_sheet.getRange("A1:H1");
      range.values = headers;
      range.format.autofitColumns();
      let header_range = log_sheet.getRange("A1:H1");
      header_range.format.fill.color = "#4472C4";
      header_range.format.font.color = "white";
      await context.sync();
      error_var = -1;
    });
  } catch (error) {
    error_var = 1;
    if (error.code == "InvalidOperationInCellEditMode") {
      // Modal for Editing Mode Error
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Load Add-In";
      document.getElementById("modalText").innerHTML = "Excel cell in Edit Mode. Please Exit Edit mode by using the Enter or Tab keys, or by selecting another cell, and then load the Add-In again.";
      myModal.show();
    }

  }
}

// Load Worksheets Dropdown 
export async function loadDropdown() {
  try {
    await Excel.run(async (context) => {
      var list = document.getElementById("sheetDropdown");

      list.length = 1;
      let sheets = context.workbook.worksheets;
      sheets.load("name");

      return context.sync().then(async function () {
        for (var k = 0; k < sheets.items.length; k++) {
          var opt = sheets.items[k].name;
          if (opt == "Sheet_Overview")
            continue
          var text = document.createTextNode(opt);
          var option = document.createElement("option");
          option.appendChild(text);
          list.appendChild(option);
        }
      });

    });
  } catch (error) {
    console.error(error);
  }
}

// Check which Option has been chosen
function radioButtonCheck() {
  const radioButtons = document.querySelectorAll('input[name="inlineRadioOptions"]');
  let selectedSize;
  for (const radioButton of radioButtons) {
    if (radioButton.checked) {
      selectedSize = radioButton.value;
      loadDropdown();
      // To IMS
      if (selectedSize === "option1") {
        document.getElementById("retrieve_select").style.display = "none";
        document.getElementById("sheet_select").style.display = "block";

        // Text box
        document.getElementById('textbox').style.display = "block";
      }
      // To Datalake
      else if (selectedSize === "option2") {
        document.getElementById("retrieve_select").style.display = "none";
        document.getElementById("sheet_select").style.display = "block";

        // Text box
        document.getElementById('textbox').style.display = "block";
      }
      // Retrieve Data
      else if (selectedSize === "option3") {
        document.getElementById("sheet_select").style.display = "none";
        document.getElementById("retrieve_select").style.display = "block";
      }
      break;
    }
  }
}


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Set Style
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Authenticate Button
    document.getElementById("signIn").onclick = signIn;

    // Instructions Button
    document.getElementById("instructions").onclick = openInstructions;

    //Signout Button
    document.getElementById("log_out").onclick = openLogOutPage;
    // Refresh Sheet Dropdown
    document.getElementById('refreshDropdown').addEventListener('click', loadDropdown);

    // Send Data Button
    document.getElementById("sendData").onclick = sendData;

    // Retrieve Data Button
    document.getElementById("run").onclick = run;

    // // Add Favourite Button
    // document.getElementById("Add").onclick= add;

    //View Favourites Button
    document.getElementById("View").onclick = view;

    // Check Radio button
    document.getElementById("inlineRadio1").addEventListener('click', radioButtonCheck);
    document.getElementById("inlineRadio2").addEventListener('click', radioButtonCheck);
    document.getElementById("inlineRadio3").addEventListener('click', radioButtonCheck);

  }
});

// Open Instructions.html
function openInstructions() {
  Excel.run(context => {
    // // sync the context to run the previous API call, and return.
    var dataToSend = {
      key1: 'value1',
      key2: 'value2'
  };
    Office.context.ui.displayDialogAsync('https://localhost:3000/instructions.html',
      // change these to your preference
      { height: 70, width: 45, promptBeforeOpen: false ,messageToParent:dataToSend},

      function (asyncResult) {

        // note _dlg is globally defined
        _dlg = asyncResult.value;

        _dlg.addEventHandler(Office.EventType.DialogMessageReceived,
          processDialogCallback);

        // Send data to the child window using the messageChild method
        _dlg.messageChild(dataToSend);
      }
    );
    return context.sync();

  });
}


// Logout Page
export async function openLogOutPage() {
  try {
    await Excel.run(async (context) => {
      // // sync the context to run the previous API call, and return.
      if (logout_url == "") {
        logout_url = "https://mingle-sso.inforcloudsuite.com/idp/startSLO.ping?TargetResource=https%3a%2f%2fmingle-portal.inforcloudsuite.com%2fetc%2fsignoutSuccess";
      }
      Office.context.ui.displayDialogAsync("https://localhost:3000/logout.html?logout_url=" + encodeURIComponent(logout_url),
        // change these to your preference
        { height: 70, width: 45, promptBeforeOpen: false },

        function (asyncResult) {

          // note _dlg is globally defined
          _dlg = asyncResult.value;

          _dlg.addEventHandler(Office.EventType.DialogMessageReceived,
            processDialogCallback);
        }
      );

      await sleep(6690);
      _dlg.close();
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Sign Out";
      document.getElementById("modalText").innerHTML = "Signed out Successfully";
      myModal.show();
      restPage()
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}


// Sign In Button Click
function signIn() {
  Excel.run(context => {

    // Change color of Authenticate Button
    document.getElementById('signIn').classList.remove('btn-success');
    document.getElementById('signIn').classList.remove('btn-danger');
    document.getElementById('signIn').classList.add('btn-primary');

    // Set Tenant Name to h6 tag
    document.getElementById('tenant_name').innerHTML = "";

    // Hide all the option when Authenticate button is pressed
    document.getElementById("option_list").style.display = "none";
    document.getElementById("sheet_select").style.display = "none";
    document.getElementById("retrieve_select").style.display = "none";

    // // sync the context to run the previous API call, and return.
    Office.context.ui.displayDialogAsync('https://localhost:3000/commands.html',
      // change these to your preference
      { height: 70, width: 45, promptBeforeOpen: false },

      function (asyncResult) {

        // note _dlg is globally defined
        _dlg = asyncResult.value;

        _dlg.addEventHandler(Office.EventType.DialogMessageReceived,
          processDialogCallback);

        //     const messageToDialog1 = JSON.stringify({
        //       city: "My Sheet"
        //   });
        // console.log(messageToDialog1);
        // console.log("Before Sending12");
        // _dlg.messageChild(messageToDialog1);
        // console.log("After Sending12");
      }
    );
    return context.sync();

  });
}

function view() {
  Excel.run(context => {
    let dataforviewfava={
      "ti":"DEVMRKT_DEV",//tenant,
      "iu":"https://mingle-ionapi.inforcloudsuite.com",//endpoint_url,
      "token":access_token
    }
    //let dataParameter = encodeURIComponent(JSON.stringify(dataforviewfava));
     dataforviewfava=JSON.stringify(dataforviewfava)
    let base64Data = btoa(dataforviewfava); 
    console.log(base64Data)
    //let compressedData = pako.deflate(base64Data, { to: 'string' });
    let viewfavurl = `https://localhost:3000/viewfavs.html?viewpage_id=7647524R7o6h5i9t8r9o90677n8767i7687&version=2s7u8t6c5h9i7t843&page_id=M123a234n3i3s43a34n45t45o543s435h&at=1l2a5h3a6r348ifjhjirhfureighjkerhgijerhguihreuigerjlgbuiritu3498579028590y8hjdgfifksjdfjffskdjfkjsb&data=${base64Data}`;
    Office.context.ui.displayDialogAsync(viewfavurl,
      // change these to your preference
      { height: 90, width: 60, promptBeforeOpen: false },
      function (asyncResult) {

        // note _dlg is globally defined
      _dlg = asyncResult.value;
      //   var dataToSend = {
      //     key1: 'value1',
      //     key2: 'value2'
      // };

        _dlg.addEventHandler(Office.EventType.DialogMessageReceived,
        processDialogCallback);
        // _dlg.postMessage(dataToSend);
      }
    );
    return context.sync();

  });

  document.getElementById("View").blur();
}
function restPage(){
      // Change color of Authenticate Button
      document.getElementById('signIn').classList.remove('btn-success');
      document.getElementById('signIn').classList.remove('btn-danger');
      document.getElementById('signIn').classList.add('btn-primary');

      // Set Tenant Name to h6 tag
      document.getElementById('tenant_name').innerHTML = "";

      // Hide all the option when Authenticate button is pressed
      document.getElementById("option_list").style.display = "none";
      document.getElementById("sheet_select").style.display = "none";
      document.getElementById("retrieve_select").style.display = "none";
      window.location.reload()
}
// Process Message received from Dialog 
async function processDialogCallback(arg) {
  var messageFromDialog = JSON.parse(arg.message);
  if (messageFromDialog.messageType === "token") {
    access_token = messageFromDialog.access_token;
    //console.log(access_token)
    //console.log(messageFromDialog)
    let expires_in   = messageFromDialog.time*1000;
    //console.log(messageFromDialog.time)
    //console.log(messageFromDialog);
    if (typeof (access_token) == 'undefined') {
      _dlg.close();
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Sign In";
      document.getElementById("modalText").innerHTML = "Failed to Sign In.";
      myModal.show();
      // Change color of Authenticate Button
      document.getElementById('signIn').classList.remove('btn-primary');
      document.getElementById('signIn').classList.add('btn-danger');
    }
    else {
      _dlg.close();
      // Change color of Authenticate Button
      document.getElementById('signIn').classList.remove('btn-primary');
      document.getElementById('signIn').classList.add('btn-success');
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      setTimeout(async () => {
        document.getElementById("modalHeading").innerHTML = "token expired";
        document.getElementById("modalText").innerHTML = "token exceded it time limit";
        myModal.show();
        restPage()
        console.log(parseInt(expires_in)*1000+typeof(parseInt(expires_in)*1000))
      }, expires_in);

      //execute API here...
      var flag_df = 0;           //var for datafabric access
      var flag_ims = 0;       //var for ims
      var flag_api = 0;         //var for api
      document.getElementById('inlineRadio1').disabled = false; // To get back to same status when a user signs out and signs in again
      document.getElementById('inlineRadio2').disabled = false;
      document.getElementById('inlineRadio3').disabled = false;
      document.getElementById('inlineRadio1').checked = false;
      document.getElementById('inlineRadio2').checked = false;
      document.getElementById('inlineRadio3').checked = false;
      var x = document.getElementById("To_IMS");
      x.title = "";
      x.style.color = '#000000';

      var y = document.getElementById("To_DL");
      y.title = "";
      y.style.color = '#000000';

      var z = document.getElementById("From_DL");
      z.title = "";
      z.style.color = '#000000';
      var result = await getPermissionsList();
      //console.log(result);

      console.log(result.response.userlist[0].groups);
      if (result.response.userlist[0].groups.find(o => o.display == 'IONAPI-User')) {
        //console.log("Found API !!");
        flag_api = 1;
      }
      else {
        //console.log("notfound API");
      }

      if (result.response.userlist[0].groups.find(p => p.display == 'IONDeskAdmin')) {
        //console.log("Found IMS!!");
        flag_ims = 1;
      }
      else {
        //console.log("notfound IMS");
      }

      if (result.response.userlist[0].groups.find(q => q.display == 'DATAFABRIC-SuperAdmin')) {
        //console.log("Found DF!!");
        flag_df = 1;
      }
      else {
        //console.log("notfound DF");
      }

      // console.log("Flag API " + flag_api)
      // console.log("Flag IMS " + flag_ims)
      // console.log("Flag DF " + flag_df)

      // Display send and recieve buttons and make checked attribute as false
      if (flag_df == 1 && flag_ims == 1) {
        //console.log("Entered Loop All");
        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Available permissions";
        document.getElementById("modalText").innerHTML = "Allowed Operations - All";
        myModal.show();
        var opts = document.getElementsByClassName('form-check');
        document.getElementById('inlineRadio1').checked = false;
        document.getElementById('inlineRadio2').checked = false;
        document.getElementById('inlineRadio3').checked = false;
        opts[0].style.display = 'block';

      }
      else if (flag_df == 1 && flag_ims == 0) {
        //console.log("Entered Loop DL");
        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Available permissions";
        document.getElementById("modalText").innerHTML = "To Data Lake and From Data Lake.Please contact Mingle Administrator for required Security Roles.";
        myModal.show();
        var opts = document.getElementsByClassName('form-check');
        document.getElementById('inlineRadio1').disabled = true;
        x = document.getElementById("To_IMS");
        x.title = "No permission to access this option";
        x.style.color = '#808080';
        document.getElementById('inlineRadio2').checked = false;
        document.getElementById('inlineRadio3').checked = false;
        opts[0].style.display = 'block';

      }
      else if (flag_df == 0 && flag_ims == 1) {
        //console.log("Entered Loop IMS ");
        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Available permissions";
        document.getElementById("modalText").innerHTML = "To IMS.Please contact Mingle Administrator for required Security Roles.";
        myModal.show();
        var opts = document.getElementsByClassName('form-check');
        document.getElementById('inlineRadio1').checked = false;
        document.getElementById('inlineRadio2').disabled = true;
        x = document.getElementById("To_DL");
        x.title = "No permission to access this option";
        x.style.color = '#808080';
        document.getElementById('inlineRadio3').disabled = true;
        y = document.getElementById("From_DL");
        y.title = "No permission to access this option";
        y.style.color = '#808080';

        opts[0].style.display = 'block';

      }

      else {

        //console.log("Final loop");
        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Available permissions";
        document.getElementById("modalText").innerHTML = "None.Please contact Mingle Administrator for required Security Roles.";
        myModal.show();
        var opts = document.getElementsByClassName('form-check');
        document.getElementById('inlineRadio1').disabled = true;
        x = document.getElementById("To_IMS");
        x.title = "No permission to access this option";
        x.style.color = '#808080';
        document.getElementById('inlineRadio2').disabled = true;
        y = document.getElementById("To_DL");
        y.title = "No permission to access this option";
        y.style.color = '#808080';
        document.getElementById('inlineRadio3').disabled = true;
        z = document.getElementById("From_DL");
        z.title = "No permission to access this option";
        z.style.color = '#808080';
        opts[0].style.display = 'block';
      }

      document.getElementById('tenant_name').innerHTML = `Tenant: ${tenant}&emsp;&nbsp`;

    }
  }

  else if (messageFromDialog.messageType === "tenant") {
    tenant = messageFromDialog.tenant_name;
    endpoint_url = messageFromDialog.endpoint_url;
    logout_url = messageFromDialog.logout_url + "/idp/startSLO.ping";
    //console.log(logout_url);
  }
  else if (messageFromDialog.messageType === "UserQuery") {
    _dlg.close();
    var UserQuery = messageFromDialog.query;
    console.log("Received query", UserQuery);
    document.getElementById("fname").value = UserQuery;
  }
  else if(messageFromDialog.messageType === "queryname"){
    var UserQuery = messageFromDialog.query;
    console.log("Received query", UserQuery);
    document.getElementById("fname").value = UserQuery;
  }
}

// Add Logs in Sheet Overview
export async function logSheet(sheet_name, schema_name, size_sheet, no_of_rows, date_time, currTime, status, error_msg, sheet_color) {
  try {
    await Excel.run(async (context) => {
      console.log(sheet_name);
      // Add logs in Sheet_Overview
      let sheets = context.workbook.worksheets;
      let log_sheet = sheets.getItem("Sheet_Overview");

      let range = log_sheet.getUsedRange();
      range.load("values");
      await context.sync();

      var sheetOverview_row = (range.values).length + 1;
      let row_range = log_sheet.getRange(`A${sheetOverview_row}:H${sheetOverview_row}`);
      if (no_of_rows == 1)
        no_of_rows = 0; // Used to assign number of rows if sheet is empty
      row_range.values = [
        [sheet_name, schema_name, size_sheet, no_of_rows, date_time, currTime, status, error_msg],
      ];
      row_range.format.autofitColumns();

      //To set color to the error messsages in Sheet_Overview
      let color_range = log_sheet.getRange(`G${sheetOverview_row}:H${sheetOverview_row}`);
      color_range.format.font.color = sheet_color;

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

// Check if schema exists in the tenant or not
async function getObjectList(schema_name) {
  try {
    var myHeaders = new Headers();
    myHeaders.append("Authorization", `Bearer ${access_token}`);

    var requestOptions = {
      method: 'GET',
      headers: myHeaders,
      redirect: 'follow'
    };
    console.log(schema_name);
    const response = await fetch(`${endpoint_url}/${tenant}/IONSERVICES/datacatalog/v1/object/list?name=${schema_name}`, requestOptions);
    const obj = await response.json();
    console.log(obj)
    //const objects = obj.objects[0];
    //console.log(count);
    return obj;
  } catch (error) {
    console.log(error);
  }
}

async function getPermissionsList() {
  try {
    var myHeaders = new Headers();
    myHeaders.append("Authorization", `Bearer ${access_token}`);

    var requestOptions = {
      method: 'GET',
      headers: myHeaders,
      redirect: 'follow'
    };

    const response = await fetch(`${endpoint_url}/${tenant}/ifsservice/usermgt/v2/users/me`, requestOptions);
    const obj = await response.json();
    //const objects = obj.objects[0];
    //console.log(count);
    return obj;
  } catch (error) {
    console.log(error);
  }
}

// Extract the final data and send it
export async function extractAndSend(json_text, schema_name, sheet_name) {
  try {
    await Excel.run(async (context) => {
      console.log(json_text);
      var empty_str = Papa.unparse(json_text.slice(1));
      console.log(empty_str);
      no_of_rows = json_text.length - 1;
      if (empty_str == "") {
        color = "Red";
        await logSheet(sheet_name, "-", 0, 0, new Date().toLocaleDateString(), new Date().toLocaleTimeString(), "", "Data Not Found", color);
        return;
      }
      // Store Log Sheet Details
      //var size_sheet = byteCount(empty_str);
      var size_sheet = byteSize(empty_str);
      var date_time = "";
      var status = "";
      var error_msg = "";

      const radioButtons = document.querySelectorAll('input[name="inlineRadioOptions"]');
      let selectedSize;
      for (const radioButton of radioButtons) {
        if (radioButton.checked) {
          selectedSize = radioButton.value;
          // Send via IMS
          if (selectedSize === "option1") {

            //Size of Data Check
            if (size_sheet > 5000000) {
              date_time = new Date().toLocaleDateString();
              error_msg = "Too Large File cannot upload";
              color = "Red";
            }

            else {
              var response = await getObjectList(schema_name);
              var schema_obj = response.objects.find(({ name }) => name === schema_name);
              if (schema_obj != undefined) {
                if (schema_obj.type == "JSON" && schema_obj.subType == undefined) {
                  var empty_str = Papa.parse(empty_str, {
                    delimiter: "",	// auto-detect
                    newline: "",	// auto-detect
                    quoteChar: '"',
                    escapeChar: '"',
                    header: true,
                    transformHeader: undefined,
                    dynamicTyping: false,
                    preview: 0,
                    encoding: "",
                    worker: false,
                    comments: false,
                    step: undefined,
                    complete: undefined,
                    error: undefined,
                    download: false,
                    downloadRequestHeaders: undefined,
                    downloadRequestBody: undefined,
                    skipEmptyLines: false,
                    chunk: undefined,
                    chunkSize: undefined,
                    fastMode: undefined,
                    beforeFirstChunk: undefined,
                    withCredentials: undefined,
                    transform: undefined,
                    delimitersToGuess: [',', '\t', '|', ';', Papa.RECORD_SEP, Papa.UNIT_SEP]
                  }).data;
                  //console.log(JSON.stringify(empty_str));
                  empty_str = JSON.stringify(empty_str);
                }

                else if (schema_obj.type == "JSON" && schema_obj.subType == "JSONStream") {
                  // Perform Pako Deflate
                  var empty_str = Papa.parse(empty_str, {
                    delimiter: "",	// auto-detect
                    newline: "",	// auto-detect
                    quoteChar: '"',
                    escapeChar: '"',
                    header: true,
                    transformHeader: undefined,
                    dynamicTyping: false,
                    preview: 0,
                    encoding: "",
                    worker: false,
                    comments: false,
                    step: undefined,
                    complete: undefined,
                    error: undefined,
                    download: false,
                    downloadRequestHeaders: undefined,
                    downloadRequestBody: undefined,
                    skipEmptyLines: false,
                    chunk: undefined,
                    chunkSize: undefined,
                    fastMode: undefined,
                    beforeFirstChunk: undefined,
                    withCredentials: undefined,
                    transform: undefined,
                    delimitersToGuess: [',', '\t', '|', ';', Papa.RECORD_SEP, Papa.UNIT_SEP]
                  }).data;
                  //console.log(JSON.stringify(empty_str));

                  empty_str = empty_str.map(JSON.stringify).join('\n');
                }
              }

              var data = JSON.stringify({
                "documentName": schema_name,
                "messageId": schema_name + Math.floor(Math.random() * 1000001).toString(),
                "fromLogicalId": `lid://${lid}`,
                "toLogicalId": "lid://default",
                "document": {
                  "value": empty_str,
                  "encoding": "NONE",
                  "characterSet": "UTF-8"
                }
              });

              var config = {
                method: 'post',
                url: `https://mingle-ionapi.inforcloudsuite.com/DEVMRKT_DEV/CustomerApi/EXCELWrapperAPI/v2/message`,
                headers: {
                  'Authorization': `Bearer ${access_token}`,
                  'Content-Type': 'application/json',
                  'cache-control': 'no-cache'
                },
                data: data
              };

              date_time = new Date().toLocaleDateString();
              var axios = require('axios');
              color = "Green";
              var result = await axios(config).catch(function (error) {
                if (error.response) {
                  status = error.response.status;
                  color = "Red";
                  if (status == 401)
                    error_msg = error.response.data["error"];
                  else
                    error_msg = error.response.data["errors"];
                }
              });

              if (result !== undefined) {
                status = `${result.data["code"]}. ${result.data["message"]}`;
                error_msg = "";
              }
            }

            // Add logs in Sheet_Overview
            await logSheet(sheet_name, schema_name, size_sheet, no_of_rows, date_time, new Date().toLocaleTimeString(), status, error_msg, color);
          }
          // Send to DataLake
          else if (selectedSize === "option2") {
            var response = await getObjectList(schema_name);
            //console.log(response);
            if (response.objects.find(({ name }) => name === schema_name) == undefined) {
              color = "Red";
              date_time = new Date().toLocaleDateString();
              // Add logs in Sheet_Overview
              await logSheet(sheet_name, schema_name, size_sheet, no_of_rows, date_time, new Date().toLocaleTimeString(), "Data Not Sent.", "Object Schema Does not Exist", color);
            }

            else {
              console.log("Entered proper loop");
              const pako = require('pako');
              var axios = require('axios');
              var FormData = require('form-data');
              var dataToUpload;
              var schema_obj = response.objects.find(({ name }) => name === schema_name);;


              if (schema_obj.type == "JSON" && schema_obj.subType == undefined) {
                // Perform Pako Deflate
                var empty_str = Papa.parse(empty_str, {
                  delimiter: "",	// auto-detect
                  newline: "",	// auto-detect
                  quoteChar: '"',
                  escapeChar: '"',
                  header: true,
                  transformHeader: undefined,
                  dynamicTyping: false,
                  preview: 0,
                  encoding: "",
                  worker: false,
                  comments: false,
                  step: undefined,
                  complete: undefined,
                  error: undefined,
                  download: false,
                  downloadRequestHeaders: undefined,
                  downloadRequestBody: undefined,
                  skipEmptyLines: false,
                  chunk: undefined,
                  chunkSize: undefined,
                  fastMode: undefined,
                  beforeFirstChunk: undefined,
                  withCredentials: undefined,
                  transform: undefined,
                  delimitersToGuess: [',', '\t', '|', ';', Papa.RECORD_SEP, Papa.UNIT_SEP]
                }).data;
                //console.log(JSON.stringify(empty_str));

                var fileAsArray = pako.deflate(JSON.stringify(empty_str));
                const compressedFile = fileAsArray.buffer;
                dataToUpload = new Blob([compressedFile], { type: 'application/json' });
              }

              else if (schema_obj.type == "JSON" && schema_obj.subType == "JSONStream") {
                // Perform Pako Deflate
                var empty_str = Papa.parse(empty_str, {
                  delimiter: "",	// auto-detect
                  newline: "",	// auto-detect
                  quoteChar: '"',
                  escapeChar: '"',
                  header: true,
                  transformHeader: undefined,
                  dynamicTyping: false,
                  preview: 0,
                  encoding: "",
                  worker: false,
                  comments: false,
                  step: undefined,
                  complete: undefined,
                  error: undefined,
                  download: false,
                  downloadRequestHeaders: undefined,
                  downloadRequestBody: undefined,
                  skipEmptyLines: false,
                  chunk: undefined,
                  chunkSize: undefined,
                  fastMode: undefined,
                  beforeFirstChunk: undefined,
                  withCredentials: undefined,
                  transform: undefined,
                  delimitersToGuess: [',', '\t', '|', ';', Papa.RECORD_SEP, Papa.UNIT_SEP]
                }).data;
                //console.log(JSON.stringify(empty_str));

                empty_str = empty_str.map(JSON.stringify).join('\n');
                var fileAsArray = pako.deflate(empty_str);
                const compressedFile = fileAsArray.buffer;
                dataToUpload = new Blob([compressedFile], { type: 'application/json' });
              }

              else {
                var fileAsArray = pako.deflate(empty_str, { to: 'string' });
                const compressedFile = fileAsArray.buffer;
                dataToUpload = new Blob([compressedFile], { type: 'text/csv;charset=utf-8' });
              }

              color = "Green";
              var data = new FormData();
              data.append('dl_document_name', schema_name);
              data.append('dl_from_logical_id', `lid://${lid}`);
              data.append('file', dataToUpload);
              var config = {
                method: 'post',
                url: `${endpoint_url}/${tenant}/DATAFABRIC/ingestion/v1/dataobjects`,
                headers: {
                  'Authorization': `Bearer ${access_token}`
                },
                data: data
              };

              date_time = new Date().toLocaleDateString();
              var axios = require('axios');
              var result = await axios(config).catch(function (error) {
                color = "Red";
                console.error(error);
                if (error.response) {
                  status = error.response.status;
                  if (status == 401)
                    error_msg = error.response.data["error"];
                  else if (status = 400) {
                    error_msg = error.response.data["errors"][0].message + " (dl_from_logical_id refers to the Logical ID being entered while sending the data)";
                  }
                  else
                    error_msg = error.response.data["errors"];
                }
              });
              if (result !== undefined) {
                status = `${result.status}. Published Successfully`;
                error_msg = "";
              }
              if (dataToUpload.size > 5000000) {
                error_msg += "Warning:The compressed file is above 5MB";
                color = "Orange";
              }

              await logSheet(sheet_name, schema_name, size_sheet, no_of_rows, date_time, new Date().toLocaleTimeString(), status, error_msg, color);

            }
          }
          break;
        }
      }
    });
  }
  catch (error) {
    console.error(error);
  }
}

// Split the Sheet Data into equal chunks
export async function splitDataToChunks(name) {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem(name);
      sheet.activate();
      let range = sheet.getUsedRange();
      range.load("address");
      await context.sync();

      // Extract rows and column from range address
      var range_str = range.address;
      range_str = range_str.slice(range_str.indexOf("!") + 1);


      var arr = range_str.split(':');
      //Get start Row and Column
      var startCol = arr[0].replace(/[0-9]/g, '');
      var startRow = arr[0].replace(/\D/g, '');

      // If Data is Empty
      if ((arr.length == 1) && (startCol + startRow == "A1")) {
        color = "Red";
        await logSheet(name, "-", 0, 0, new Date().toLocaleDateString(), new Date().toLocaleTimeString(), "", "Data Not Found", color);
      }

      else {
        // Get End Row and Column
        var endCol = arr[1].replace(/[0-9]/g, '');
        var endRow = arr[1].replace(/\D/g, '');
        startRow = parseInt(startRow);

        var number = endRow;
        var chunk_array = new Array(Math.floor(number / 10000)).fill(10000).concat(number % 10000);

        // Var to Hold Data
        var json_text = [];
        var rowend = 0;
        for (var j = 0; j < chunk_array.length; j++) {
          rowend += chunk_array[j];
          let end = endCol + rowend;
          end = `${startCol}${startRow}: ${end}`;
          let range = sheet.getRange(end);
          range.load("text");
          await context.sync();
          console.log(range.text);
          json_text.push(range.text);
          startRow = rowend + 1;
        }

        var finalDataToSend = [];

        //console.log(json_text);
        for (var i = 0; i < json_text.length; i++) {
          for (var j = 0; j < json_text[i].length; j++) {
            finalDataToSend.push(json_text[i][j]);
          }
        }

        await extractAndSend(finalDataToSend, finalDataToSend[0][0], name);
      }
    });
  }
  catch (error) {
    console.error(error);
  }
}

// Send Data Button Click
export async function sendData() {
  try {
    await Excel.run(async (context) => {

      // Check logical ID
      lid = document.getElementById("lid").value;

      const radioButtons = document.querySelectorAll('input[name="inlineRadioOptions"]');
      let selectedSize;
      for (const radioButton of radioButtons) {
        if (radioButton.checked) {
          selectedSize = radioButton.value;
          if (lid == "") {
            var myModal = new bootstrap.Modal(document.getElementById("myModal"));
            document.getElementById("modalHeading").innerHTML = "Send Data";
            document.getElementById("modalText").innerHTML = "Please Enter Logical ID";
            myModal.show();
            return;
          }
        }
      }
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Send Data";
      document.getElementById("modalText").innerHTML = "Processing Data. Please Wait";
      myModal.show();
      var select = document.getElementById('sheetDropdown');
      var text = select.options[select.selectedIndex].text;

      if (text == "ALL") {
        let sheets = context.workbook.worksheets;
        sheets.load("name");

        // Check is Log Sheet exists or not
        if (error_var == 0) {
          addSheet();
        }

        return context.sync().then(async function () {

          for (var k = 0; k < sheets.items.length - 1; k++) {
            let sheet = context.workbook.worksheets.getItem(sheets.items[k].name);
            sheet.activate();
            sheet.load("name");
            await context.sync();

            if (sheet.name === 'Sheet_Overview') {
              break;
            }

            // Split the data into chunks
            await splitDataToChunks(sheet.name);
          }

          // If All Sheets being sent, then make Sheet_Overview active
          let sheet = context.workbook.worksheets.getItem('Sheet_Overview');
          sheet.activate();
          document.getElementById("modalText").innerHTML = "Data Processed. Please check Sheet_Overview for more details.";
        }).catch(e => {
          console.log(e);
        });
      }

      else {
        // Check is Log Sheet exists or not
        if (error_var == 0) {
          addSheet();
        }

        let sheet = context.workbook.worksheets.getItem(text);
        sheet.activate();

        sheet.load("name");
        await context.sync();

        // Split the data into chunks
        await splitDataToChunks(sheet.name);

        document.getElementById("modalText").innerHTML = "Data Processed. Please check Sheet_Overview for more details.";

        // If Single Sheet is being sent, then make Sheet_Overview active
        let sheet_overview = context.workbook.worksheets.getItem('Sheet_Overview');
        sheet_overview.activate();
      }

    });
  } catch (error) {
    console.error(error);
  }
}

//function byteCount(s) {
//   return encodeURI(s).split(/%..|./).length - 1;
// }
const byteSize = str => new Blob([str]).size;


// Data Retrieval Process

// Get Query ID
async function getQueryId(fname) {
  try {
    var myHeaders = new Headers();
    myHeaders.append("Authorization", `Bearer ${access_token}`);
    myHeaders.append("Content-Type", "text/plain");

    var raw = fname;

    var requestOptions = {
      method: 'POST',
      headers: myHeaders,
      body: raw,
      redirect: 'follow'
    };
    const response = await fetch(`${endpoint_url}/${tenant}/DATAFABRIC/compass/v2/jobs/`, requestOptions);
    const obj = await response.json();
    return obj;
  } catch (error) {
    console.error(error);
  }
}

// Check Status of Query
async function checkStatus(queryId) {
  try {
    var myHeaders = new Headers();
    myHeaders.append("Authorization", `Bearer ${access_token}`);

    var requestOptions = {
      method: 'GET',
      headers: myHeaders,
      redirect: 'follow'
    };

    const response = await fetch(`${endpoint_url}/${tenant}/DATAFABRIC/compass/v2/jobs/${queryId}/status/`, requestOptions);
    const obj = await response.json();
    return obj;
  } catch (error) {
    console.error(error);
  }
}

// Get data from Result
async function getResult(queryId,i,limivalcheck) {
  try {
    var myHeaders = new Headers();
    let resultURL;
    myHeaders.append("Authorization", `Bearer ${access_token}`);

    var requestOptions = {
      method: 'GET',
      headers: myHeaders,
      redirect: 'follow'
    };

    // Read Limit and Offset values
    if(limivalcheck<100000){
      resultURL = `${endpoint_url}/${tenant}/DATAFABRIC/compass/v2/jobs/${queryId}/result?limit=${limivalcheck}&offset=${i}`;
    }
    else{
    resultURL = `${endpoint_url}/${tenant}/DATAFABRIC/compass/v2/jobs/${queryId}/result?limit=100000&offset=${i}`;
    }

    const response = await fetch(resultURL, requestOptions);
    statusMessage = response.status
    // const isSuccessful = response.ok;
    // if (isSuccessful) {
    // // do something
    // offset=offset+limit
    // limit=limit-1
    // console.log(offset)
    // console.log(limit)
    // }
    const obj = await response.text();
    return obj;
  } catch (error) {
    console.log(error);
  }
}

// Rename worksheet with provided Sheet name 
export async function renameWorksheet(sheet) {
  try {
    await Excel.run(async (context) => {
      var name = document.getElementById('sheet_name').value;
      sheet.name = name;
      await context.sync();
    });
  }
  catch (error) {
    console.error(error);
  }
}

// Get Column Name from Number
function printString(columnNumber) {
  // To store result (Excel column name)
  let columnName = [];

  while (columnNumber > 0) {
    // Find remainder
    let rem = columnNumber % 26;

    // If remainder is 0, then a
    // 'Z' must be there in output
    if (rem == 0) {
      columnName.push("Z");
      columnNumber = Math.floor(columnNumber / 26) - 1;
    }
    else // If remainder is non-zero
    {
      columnName.push(String.fromCharCode((rem - 1) + 'A'.charCodeAt(0)));
      columnNumber = Math.floor(columnNumber / 26);
    }
  }
  columnName = columnName.reverse().join("")
  return columnName;
}

// Function to cause delay for every iteration
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function add() {
  // console.log("test");

  let query = document.getElementById('fname').value;
  localStorage.setItem(`Query_${querynum}`, query);
  querynum++;
}
// Get response data
var statusMessage 

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      var name = document.getElementById('sheet_name').value;
      if (name.length > 31) {
        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Retrieve Data";
        document.getElementById("modalText").innerHTML = "Please provide the sheet name which has less than 31 characters";
        myModal.show();
        return;
      }

      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Retrieve Data";
      document.getElementById("modalText").innerHTML = "Retrieving Data. Please Wait";
      myModal.show();
      let fname = document.getElementById('fname').value;
      //console.log(fname)
      let sheet = context.workbook.worksheets.getActiveWorksheet();

      var response = await getQueryId(fname);
      var queryId = response.queryId;
      //console.log(queryId);
      var count = 60;

      const apiResponseTimeStart = Date.now();
      while (count >= 0) {
        response = await checkStatus(queryId);
        if (response.status == "FINISHED") {
        let limit = document.getElementById('limit').value;
        let offset = document.getElementById('offset').value;
        var limivalcheck = limit
        let i = parseInt(offset) || 0
        console.log(i)
        var rowno = 0;
        var rowstart = 1;
        var rowend = 0;
        //for(let i=0;i<=limit;i=i+10)
        const apiResponseTimeEnd = Date.now();
        console.log(`API Response Retrieval Time: ${apiResponseTimeEnd - apiResponseTimeStart} ms`);
        document.getElementById("modalText").innerHTML = "Retrieved Data From ION. Loading Data into Excel Sheet";
        
        var sheetDataLoadTimeStart = Date.now();
        try{
        var tablehead=0
        var breaktheloop=0
        while(limivalcheck>=0){
          //console.log(rowstart)
          try{
              response = await getResult(queryId,i,limivalcheck);
              if(response.data!="" && statusMessage==200)
              {
              i=i+100000
              tablehead=tablehead+1
              limivalcheck=limivalcheck-100000
              console.log(limivalcheck)
              }
              else if(statusMessage!=200){
                breaktheloop=breaktheloop+1
                if(breaktheloop>=2){
                  break
                } 
              }

              
              //console.log(response + "responce")
              let data = Papa.parse(response).data;
              if(tablehead>1){
                data.shift()
              }
              //console.log(data + "data")
              //console.log(data[0] + "data")
              //var size_sheet = byteCount(response); // Get size of data
              var size_sheet = byteSize(response)+size_sheet;
              console.log(size_sheet + "size_sheet")
              let columnno = 0;

              //data.pop();
              console.log(data.pop() + "data.pop()")
              rowno = rowno+data.length;
              columnno = data[0].length;

              let columnname = printString(columnno);

              // For Loop to Split the Retrieval Process
              if (data.length <= 1) {
                let end = columnname + rowno;
                end = `A1: ${end}`;
                let range = sheet.getRange(end);
                range.values = data;
                range.format.autofitColumns();
                await context.sync();
              }

              else {
                var index = 0;
                var chunk_size = 10000;
                var arrayLength = data.length;
                var tempArray = [];

                
                columnno = data[0].length;

                let columnname = printString(columnno);

                for (index = 0; index < arrayLength; index += chunk_size) {
                  console.log(data.slice(index, index + chunk_size))
                  tempArray.push(data.slice(index, index + chunk_size));
                }

                for (var j = 0; j < tempArray.length; j++) {
                  rowend += tempArray[j].length;
                  let end = columnname + rowend;
                  end = `A${rowstart}: ${end}`;
                  let range = sheet.getRange(end);
                  range.values = tempArray[j];
                  range.format.autofitColumns();
                  await context.sync();
                  rowstart = rowend + 1;
                }
              }
            }
            catch(error){
              console.log(error)
              break
            }

        }
        }
        catch(error){
          console.log(error)
        }
        i=0
        const sheetDataLoadTimeEnd = Date.now();

        // rename only if name exists
        if(name)
          await renameWorksheet(sheet);

        document.getElementById("modalText").innerHTML = "Retrieved Data Successfully.";
        console.log(`Sheet Data Load Time: ${sheetDataLoadTimeEnd - sheetDataLoadTimeStart} ms`);

        // Log details into Sheet_Overview
        sheet.load('name');
        await context.sync();
        var sheet_name = sheet.name;
        var date_time = new Date().toLocaleDateString();
        var status = "";
        var apiResponseTimeTotal = apiResponseTimeEnd - apiResponseTimeStart;
        var sheetDataLoadTimeTotal = sheetDataLoadTimeEnd - sheetDataLoadTimeStart;
        var totalLoadTime = sheetDataLoadTimeEnd - apiResponseTimeStart;

        status = status + "Retrieved Data Successfully. Query Id: " + queryId
          + ".\nAPI Response Retrieval Time: " + (Math.floor(apiResponseTimeTotal / 60000) + ":" + (((apiResponseTimeTotal % 60000) / 1000).toFixed(0) < 10 ? '0' : '') + ((apiResponseTimeTotal % 60000) / 1000).toFixed(0)) + ".\nSheet Data LoadTime: " + (Math.floor(sheetDataLoadTimeTotal / 60000) + ":" + (((sheetDataLoadTimeTotal % 60000) / 1000).toFixed(0) < 10 ? '0' : '') + ((sheetDataLoadTimeTotal % 60000) / 1000).toFixed(0)) + ".\nTotal Time for Retrieval: " + (Math.floor(totalLoadTime / 60000) + ":" + (((totalLoadTime % 60000) / 1000).toFixed(0) < 10 ? '0' : '') + ((totalLoadTime % 60000) / 1000).toFixed(0)) + ".";
        var error_msg = "";
        color = "Green";
        await logSheet(sheet_name, "-", size_sheet, rowno, date_time, new Date().toLocaleTimeString(), status, error_msg, color);
          break;
        }

        else if (response.status == "FAILED") {
          var date_time = new Date().toLocaleDateString();
          var status = response.status;
          status = status + ". Couldn't retrieve the data please check the query.";
          var error_msg = "Couldn't retrieve the data please check the query. Query Id: " + queryId;
          color = "Red";
          await logSheet("-", "-", 0, 0, date_time, new Date().toLocaleTimeString(), status, error_msg, color);
          document.getElementById("modalText").innerHTML = "Could not retrieve data ,please check Sheet_Overview for more details";
          break;
        }
        else if (count == 1 && response.status == "RUNNING") {
          var date_time = new Date().toLocaleDateString();
          var status = response.status;
          var error_msg = "Please Re-Run the Query. Query Id: " + queryId;
          color = "Orange";
          await logSheet("-", "-", 0, 0, date_time, new Date().toLocaleTimeString(), status, error_msg, color);
          document.getElementById("modalText").innerHTML = "Please Re-Run the Query.";
          break;
        }
        count--;
        await sleep(5000);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

