/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

var code;
var auth_obj;
var token;
var file_count = 0;
var flag = 0; // Check for same ionapi file

// OnLoad Extract Code and URL
window.addEventListener('load', () => {
  // Load Dropdown OnLoad
  var list = document.getElementById("ionAPIDropdown");

  for (var k = localStorage.length - 1; k >= 1; k--) {
    var opt = localStorage.getItem(`ionAPI${k}`);
    var text = document.createTextNode(JSON.parse(opt).cn);
    var option = document.createElement("option");
    option.appendChild(text);
    list.appendChild(option);
  }

  // Call API
  const urlParams = new URLSearchParams(location.search);

  for (const [key, value] of urlParams) {
    if (key == "code") {
      console.log(`${key}:${value}`);
      code = value;
      //Office.context.ui.messageParent(code);
      // // Call API for access token
      // var axios = require('axios')
      // var qs = require('qs');
      // var data = qs.stringify({
      //   'grant_type': 'authorization_code',
      //   'client_id': 'MANDALA_DEM~5Zk2gCbQ01r1M58ba9ZR7UzDjjwQwSO5A53qKzAVd7o',
      //   'client_secret': 'NwNygxJQHiw3EDn7I32N3MWQ4r0NQf8aJQpmSsuURAYhWowbaK4F4XMySnnJnBqEfjiX_m-qGh3TynHTwvojug',
      //   'code': code,
      //   'redirect_uri': 'https://localhost:3000/commands.html'
      // });

      // var config = {
      //   method: 'post',
      //   url: 'https://stormy-river-96053.herokuapp.com/https://mingle-sso.inforcloudsuite.com:443/MANDALA_DEM/as/token.oauth2',
      //   headers: {
      //     'Content-Type': 'application/x-www-form-urlencoded'
      //   },
      //   data: data
      // };

      // axios(config)
      //   .then(function (response) {
      //     Office.context.ui.messageParent(response.data.access_token);
      //     //console.log(JSON.stringify(response.data));
      //   })
      //   .catch(function (error) {
      //     console.log(error);
      //   });



      // Fetch Token Generation
      const fetch = (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));
      var myHeaders = new Headers();
      myHeaders.append("Content-Type", "application/x-www-form-urlencoded");

      // Add For loop to Update Tenant details
      var urlencoded = new URLSearchParams();
      urlencoded.append("grant_type", "authorization_code");
      urlencoded.append("client_id", "MANDALA_DEM~5Zk2gCbQ01r1M58ba9ZR7UzDjjwQwSO5A53qKzAVd7o");
      urlencoded.append("client_secret", "NwNygxJQHiw3EDn7I32N3MWQ4r0NQf8aJQpmSsuURAYhWowbaK4F4XMySnnJnBqEfjiX_m-qGh3TynHTwvojug");
      urlencoded.append("code", code);
      urlencoded.append("redirect_uri", "https://localhost:3000/commands.html");

      var requestOptions = {
        method: 'POST',
        headers: myHeaders,
        body: urlencoded,
        redirect: 'follow'
      };

      const address = fetch("https://stormy-river-96053.herokuapp.com/https://mingle-sso.inforcloudsuite.com:443/MANDALA_DEM/as/token.oauth2", requestOptions)
        .then(response => response.json())
        .then((result) => {
          return result
        })
        .catch(error => console.log('error', error));

      const printAddress = async () => {
        const a = await address;
        var messageObject = { messageType: "token", access_token: a.access_token };
        var jsonMessage = JSON.stringify(messageObject);
        Office.context.ui.messageParent(jsonMessage);
      };

      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Sign In";
      document.getElementById("modalText").innerHTML = `Signed in Successfully.`;
      myModal.show();
      printAddress();
    }
  }
});

function displayPopup() {
  var name = document.getElementById('ionAPIDropdown').value;

  for (var k = localStorage.length - 1; k >= 1; k--) {
    var opt = JSON.parse(localStorage.getItem(`ionAPI${k}`));

    if (opt.cn == name) {
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Upload Profile";
      document.getElementById("modalText").innerHTML = `Profile Uploaded Successfully. Please Sign in into ${opt.ti} tenant.`;
      myModal.show();
      break;
    }
  }
}

Office.onReady((info) => {
  // If needed, Office.js is ready to be called
  if (info.host === Office.HostType.Excel) {

    //Button Event Capture
    document.getElementById("uploadBtn").addEventListener('click', openDialog);
    document.getElementById("fileid").addEventListener('change', readSingleFile, false);


    // Login to Infor
    document.getElementById('logIn').onclick = logIn;

    // Select Tag Onchange
    document.getElementById('ionAPIDropdown').onchange = displayPopup;
    // document.getElementById("modal").addEventListener('click', function() {
    //   var myModal = new bootstrap.Modal(document.getElementById("myModal"));
    //   myModal.show();
    // });
  }
});


//Upload Button Click
function openDialog() {
  document.getElementById('fileid').click();
}

function readSingleFile(e) {

  var file = e.target.files[0];
  console.log(e.value);
  console.log(file.name)
  console.log(file.name.split('.').pop());
  if (file.name.split('.').pop() == "ionapi") {
    if (!file) {
      return;
    }

    file_count += 1;
    var reader = new FileReader();
    reader.onload = function (e) {
      auth_obj = JSON.parse(e.target.result);
      console.log(auth_obj);
      if(!('ru' in auth_obj))
      {
        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Upload Profile";
        document.getElementById("modalText").innerHTML = "Please Upload '.ionapi' file of type WebApp";
        myModal.show();
        return;
      }
      // var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      // document.getElementById("modalHeading").innerHTML = "Upload Profile";
      // document.getElementById("modalText").innerHTML = `Profile Uploaded Successfully. Please Sign in into ${auth_obj.ti} tenant.`;
      // myModal.show();

      // Check Same ionapi File
      flag = 0;
      for (var k = localStorage.length - 1; k >= 1; k--) {
        var opt = JSON.parse(localStorage.getItem(`ionAPI${k}`));

        if (opt.cn == auth_obj.cn) {
          var selectElement = document.getElementById('ionAPIDropdown');
          selectElement.length = 1; // Delete all options
          localStorage.removeItem(`ionAPI${k}`);
          localStorage.setItem(`ionAPI${k}`, JSON.stringify(auth_obj));
          file_count--; //Decrement as Same File replaced
          flag = 1;

          // Send tenant details
          var messageObject = { messageType: "tenant", tenant_name: auth_obj.ti };
          var jsonMessage = JSON.stringify(messageObject);
          Office.context.ui.messageParent(jsonMessage);

          // Display Modal
          var myModal = new bootstrap.Modal(document.getElementById("myModal"));
          document.getElementById("modalHeading").innerHTML = "Upload Profile";
          document.getElementById("modalText").innerHTML = `Profile Uploaded Successfully. Please Sign in into ${auth_obj.ti} tenant.`;
          myModal.show();


          // Reset Dropdown
          var list = document.getElementById("ionAPIDropdown");

          for (var k = localStorage.length - 1; k >= 1; k--) {
            var opt = localStorage.getItem(`ionAPI${k}`);
            var text = document.createTextNode(JSON.parse(opt).cn);
            var option = document.createElement("option");
            option.appendChild(text);
            list.appendChild(option);
            if (option.text == auth_obj.cn) {
              option.selected = true;
            }
          }
          break;
        }
      }

      if (flag == 0) {
        localStorage.setItem(`ionAPI${file_count}`, JSON.stringify(auth_obj));
        // Send tenant details
        var messageObject = { messageType: "tenant", tenant_name: auth_obj.ti };
        var jsonMessage = JSON.stringify(messageObject);
        Office.context.ui.messageParent(jsonMessage);

        // Load Dropdown Values
        var list = document.getElementById("ionAPIDropdown");
        var opt = localStorage.getItem(`ionAPI${file_count}`);
        var text = document.createTextNode(JSON.parse(opt).cn);
        var option = document.createElement("option");
        option.appendChild(text);
        list.append(option);
        option.selected = true;

        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Upload Profile";
        document.getElementById("modalText").innerHTML = `Profile Uploaded Successfully. Please Sign in into ${JSON.parse(opt).ti} tenant.`;
        myModal.show();
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  }

  else {
    var myModal = new bootstrap.Modal(document.getElementById("myModal"));
    document.getElementById("modalHeading").innerHTML = "Upload Profile";
    document.getElementById("modalText").innerHTML = "Please Upload '.ionapi' file.";
    myModal.show();
    return;
  }


}





// Login Function
function logIn() {
  var name = document.getElementById('ionAPIDropdown').value;

  for (var k = localStorage.length - 1; k >= 1; k--) {
    var auth_obj = JSON.parse(localStorage.getItem(`ionAPI${k}`));

    if (auth_obj.cn == name) {
      window.location.replace(`${auth_obj['pu']}${auth_obj['oa']}?client_id=${auth_obj['ci']}&response_type=code&redirect_uri=${auth_obj['ru']}`);
      break;
    }
  }

  //console.log(`${auth_obj['pu']}${auth_obj['oa']}?client_id=${auth_obj['ci']}&response_type=code&redirect_uri=${auth_obj['ru']}`);
}



/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
