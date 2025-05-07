/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const { create } = require("core-js/core/object");

/* global document, Office */

// Office.onReady((info) => {
//     if (info.host === Office.HostType.Outlook) {
//       document.getElementById("sideload-msg").style.display = "none";
//       document.getElementById("app-body").style.display = "flex";
//       document.getElementById("run").onclick = run;
//     }
//   });

var regardingItem = null;

////////////STAGING/////////////////////////////////
// const createemailapi = "https://prod-200.westeurope.logic.azure.com:443/workflows/7fc3dd1d8348461bb773102354791678/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=9Wd5IGNofy0lJFr1u0YKqPtifAVjI5d1UxyFGQv14kk";
// const searchregardingapi = "https://prod-81.westeurope.logic.azure.com:443/workflows/1ddbbdd778ee4104991266039f724f4a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=9jnkdoeMtPv6_ficU6q2RKTxLymJQxnGDGozYUDfZpg";
// const searchmissingemailsapi = "https://prod-80.westeurope.logic.azure.com:443/workflows/45dcfb9f75d04f1a8ad03f2996ff94e8/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=hPy8mfFdYQI_hVDNpJ_DH-vfWj1gygrXTAPWQaiF9U8";

////////////LIVE/////////////////////////////////
const createemailapi = "https://prod-60.westeurope.logic.azure.com:443/workflows/d176927b3cac453e8f3c41b812655c7e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=1nlpUWi5lOdE2k9iFTx7-FTdgQooyPvmYgybyiAqNs0";
const searchregardingapi = "https://prod-150.westeurope.logic.azure.com:443/workflows/fb5e125f2eb640bf8aba86b15b9aeb03/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=GqpQnsIgYp_2hL1bz2S9EpUpxr60qQyfQSdttetQQLI";
const searchmissingemailsapi = "https://prod-73.westeurope.logic.azure.com:443/workflows/37d8b1ec35454bfcbc5ca129c06823af/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jUMDas16sL2BarEnMn4WOkP-MayqLNpJY5_-skRw1tQ";
  
function loadControls() {
  loadMissingEmails();
  document.getElementById("searchBox").addEventListener("keydown", function(event) {
    if (event.key === "Enter") {
        event.preventDefault(); // Prevent default form submission
        const value = event.target.value;
        showSuggestionsOnEnter(value); // Call the function
    }
  });
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {

     
          // Your code here
    const item = Office.context.mailbox.item;
    const dateTimeCreated = item.dateTimeCreated;
    const dateField = document.getElementById("mailtime");
    console.log("message type is " + Office.context.mailbox.item.messageType + " and item type is " + Office.context.mailbox.item.itemType);
    //const utcDateString = "2025-04-15T12:00:00Z"; // UTC time
    //const utcDateString = dateTimeCreated.format("YYYY-MM-DDTHH:mm:ssZ"); // UTC time
    //console.log(utcDateString);
    //const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone; // Get user's current time zone
    //const localTime = convertUTCToLocalTime(utcDateString, timeZone);
  
    //const dt = new Date(localTime);
  
    const formatteddate = formatDateToISO(dateTimeCreated);
  
    dateField.value = formatteddate;
    }
  });



}

function loadMissingEmails() {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      document.getElementById("run").innerHTML = "Tracking.....";
      const item = Office.context.mailbox.item;
          
      const missingemailstext = document.getElementById("missingemailstext");
        
      const myHeaders = new Headers();
      myHeaders.append("Content-Type", "application/json");
      const from = item.from;
      const to = item.to;
      const cc = item.cc;
      const raw = JSON.stringify({
        from: from,
        to: to,
        cc: cc
      });

      const requestOptions = {
        method: "POST",
        headers: myHeaders,
        body: raw,
        redirect: "follow",
      };

      fetch(
        searchmissingemailsapi,
        requestOptions
      )
      .then((response) => response.json())
      .then((result) => {
        if (result.missingemails.length > 0) {
          const missingemailstitle = document.getElementById("missingemailstitle");
          missingemailstitle.style.display = "block";
          missingemailstext.innerText = result.missingemails;
        }
        
      })
      .catch((error) => console.error(error));
    }
  });
}

 async function run() {
    /**
     * Insert your Outlook code here
     */
    document.getElementById("run").innerHTML = "Tracking.....";
    const item = Office.context.mailbox.item;
    var userProfile = Office.context.mailbox.userProfile;
        
    // Get the user's email address
    var userEmailAddress = userProfile.emailAddress;
    console.log("User's email address: " + userEmailAddress);

    let insertAt = document.getElementById("item-subject");
  
    // insertAt.appendChild(document.createElement("br"));
    // insertAt.appendChild(document.createTextNode(item.subject));
    // insertAt.appendChild(document.createElement("br"));
    // insertAt.appendChild(document.createTextNode(item.from.displayName));
    // insertAt.appendChild(document.createElement("br"));
    // insertAt.appendChild(document.createTextNode(item.from.emailAddress));
    // insertAt.appendChild(document.createElement("br"));
    // insertAt.appendChild(document.createTextNode(item.conversationId));
  
    console.log(item);
    const dateField = document.getElementById("mailtime");
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    const from = item.from;
    const to = item.to;
    const cc = item.cc;
    const subject = item.subject;
    const trackingid = item.conversationId;
    const dateTimeCreated = item.dateTimeCreated;
    const dateTimeCreatedUTC = convertLocalToUTC(dateField.value);
    console.log("dateTimeCreatedUTC: ", dateTimeCreatedUTC);
    item.body.getAsync("html", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Successfully retrieved the email body
        const raw = JSON.stringify({
          from: from,
          to: to,
          cc: cc,
          subject: subject,
          description: result.value,
          useremailaddress: userEmailAddress,
          trackingid: trackingid,
          //dateTimeCreated: dateTimeCreated.format("YYYY-MM-DDTHH:mm:ss")
          dateTimeCreated: dateTimeCreatedUTC
        });
  
        const requestOptions = {
          method: "POST",
          headers: myHeaders,
          body: raw,
          redirect: "follow",
        };
  
        fetch(
          createemailapi,
          requestOptions
        )
          .then((response) => {
            if (response.ok) {
              let label = document.createElement("b").appendChild(document.createTextNode("Email successfully created."));
              insertAt.appendChild(label);
              document.getElementById("run").innerHTML = "Track in Jumla";
              if (item.itemType === Office.MailboxEnums.ItemType.Message) {
                const categoryToCheck = "Tracked To Jumla";
                const mycategories = [categoryToCheck];
                try {
                  // Add a category
                  item.categories.addAsync(mycategories, function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                      console.log("Category added successfully.");
                    } else {
                      console.error("Failed to add category: " + result.error.message);
                    }
                  });
                } catch (error) {
                  console.error("Error: ", error);
                }
              }
            }
          })
          .then((result) => console.log(result))
          .catch((error) => {
            let label = document.createElement("b").appendChild(document.createTextNode(error));
            insertAt.appendChild(label);
            document.getElementById("run").innerHTML = "Track in Jumla";
          });
        // Do something with the email body here
      } else {
        // Handle error
        console.log("Error: ", result.error.message);
      }
    });
  
    console.log("Item: ", item);
  }
  
  function convertUTCToLocalTime(utcDateString, timeZone) {
    const utcDate = new Date(utcDateString);
    const localDate = utcDate.toLocaleString("en-US", { timeZone: timeZone });
    return localDate;
  }
  
  function formatDateToISO(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are 0-indexed
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
  
    return `${year}-${month}-${day}T${hours}:${minutes}`;
  }

  function convertLocalToUTC(localdate) {
    const date = new Date(localdate);
    const utcYear = date.getUTCFullYear();
    const utcMonth = String(date.getUTCMonth() + 1).padStart(2, '0'); // Months are 0-indexed
    const utcDay = String(date.getUTCDate()).padStart(2, '0');
    const utcHours = String(date.getUTCHours()).padStart(2, '0');
    const utcMinutes = String(date.getUTCMinutes()).padStart(2, '0');
    const utcSeconds = String(date.getUTCSeconds()).padStart(2, '0');
  
    return `${utcYear}-${utcMonth}-${utcDay}T${utcHours}:${utcMinutes}:${utcSeconds}`;
  }

  // const suggestionsList = ["Apple","Almond","Asbestos", "Banana", "Cherry", "Date", "Grapes", "Mango", "Orange", "Pineapple", "Strawberry"];
  const suggestionsList = [];

function showSuggestions(value) {
    const suggestionsDiv = document.getElementById("suggestions");
    suggestionsDiv.innerHTML = "";
    if (value.length === 0) {
        suggestionsDiv.style.display = "none";
        return;
    }
    
    const filteredSuggestions = suggestionsList.filter(item => item.toLowerCase().startsWith(value.toLowerCase()));
    
    if (filteredSuggestions.length > 0) {
        suggestionsDiv.style.display = "block";
        filteredSuggestions.forEach(suggestion => {
            const div = document.createElement("div");
            div.classList.add("suggestion-item");
            div.innerText = suggestion;
            div.onclick = () => {
                document.getElementById("searchBox").value = suggestion;
                suggestionsDiv.style.display = "none";
            };
            suggestionsDiv.appendChild(div);
        });
    } else {
        suggestionsDiv.style.display = "none";
    }
}

function clearSuggestions(value) {
  const suggestionsDiv = document.getElementById("suggestions");
  suggestionsDiv.innerHTML = "";
  suggestionsDiv.style.display = "none";
  if (value.length === 0) {
      // suggestionsDiv.style.display = "none";
      regardingItem = null;
      console.log("regardingItem: ", regardingItem);
  }
}

function showSuggestionsOnEnter(value) {
  const suggestionsDiv = document.getElementById("suggestions");
  const searchText = document.getElementById("searchText");
  searchText.style.display = "block";
  //show loading animation
  // const divload = document.createElement("div");
  // divload.classList.add("suggestion-item");
  // divload.innerText = "searching...";
  // suggestionsDiv.appendChild(divload);
  // suggestionsDiv.style.display = "block";

  const myHeaders = new Headers();
myHeaders.append("Content-Type", "application/json");

const raw = JSON.stringify({
  "input": value
});

const requestOptions = {
  method: "POST",
  headers: myHeaders,
  body: raw,
  redirect: "follow"
};

fetch(searchregardingapi, requestOptions)
  .then((response) => response.json())
  .then((result) => {
    searchText.style.display = "none";
    
    suggestionsDiv.innerHTML = "";
    if (result.length === 0) {
        suggestionsDiv.style.display = "none";
        return;
    }
    else{
        suggestionsList.length = 0; // Clear the array
        result.forEach(item => {
            suggestionsList.push(item);
        });
    }
    
    // const filteredSuggestions = suggestionsList.filter(item => item.toLowerCase().startsWith(value.toLowerCase()));
    const filteredSuggestions = suggestionsList;
    
    if (filteredSuggestions.length > 0) {
        suggestionsDiv.style.display = "block";
        filteredSuggestions.forEach(suggestion => {
            const div = document.createElement("div");
            div.classList.add("suggestion-item");
            div.innerText = suggestion.name + " (" + suggestion.recordtype + ")";
            div.onclick = () => {
                document.getElementById("searchBox").value = suggestion.name + " (" + suggestion.recordtype + ")";
                regardingItem = suggestion;
                suggestionsDiv.style.display = "none";
                console.log("regardingItem: ", regardingItem);
            };
            suggestionsDiv.appendChild(div);
        });
    } else {
        suggestionsDiv.style.display = "none";
    }

  })
  .catch((error) => console.error(error));


}


