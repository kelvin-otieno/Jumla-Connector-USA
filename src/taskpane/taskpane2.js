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
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    const from = item.from;
    const to = item.to;
    const cc = item.cc;
    const subject = item.subject;
    const trackingid = item.conversationId;
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
  