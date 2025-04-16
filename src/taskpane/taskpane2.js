/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Office.onReady((info) => {
//     if (info.host === Office.HostType.Outlook) {
//       document.getElementById("sideload-msg").style.display = "none";
//       document.getElementById("app-body").style.display = "flex";
//       document.getElementById("run").onclick = run;
//     }
//   });
  
function loadControls() {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
          // Your code here
    const item = Office.context.mailbox.item;
    const dateTimeCreated = item.dateTimeCreated;
    const dateField = document.getElementById("mailtime");
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
          "https://prod-60.westeurope.logic.azure.com:443/workflows/d176927b3cac453e8f3c41b812655c7e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=1nlpUWi5lOdE2k9iFTx7-FTdgQooyPvmYgybyiAqNs0",
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