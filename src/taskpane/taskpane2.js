/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//const { create } = require("core-js/core/object");

/* global document, Office */

// Office.onReady((info) => {
//     if (info.host === Office.HostType.Outlook) {
//       document.getElementById("sideload-msg").style.display = "none";
//       document.getElementById("app-body").style.display = "flex";
//       document.getElementById("run").onclick = run;
//     }
//   });

var regardingItem = null;
var regardingItemOpp = null;
var attachments = [];


////////////STAGING/////////////////////////////////
// const createemailapi = "https://prod-200.westeurope.logic.azure.com:443/workflows/7fc3dd1d8348461bb773102354791678/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=9Wd5IGNofy0lJFr1u0YKqPtifAVjI5d1UxyFGQv14kk";
// const searchregardingapi = "https://prod-81.westeurope.logic.azure.com:443/workflows/1ddbbdd778ee4104991266039f724f4a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=9jnkdoeMtPv6_ficU6q2RKTxLymJQxnGDGozYUDfZpg";
// const searchmissingemailsapi = "https://prod-80.westeurope.logic.azure.com:443/workflows/45dcfb9f75d04f1a8ad03f2996ff94e8/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=hPy8mfFdYQI_hVDNpJ_DH-vfWj1gygrXTAPWQaiF9U8";

////////////LIVE/////////////////////////////////
let createemailapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d176927b3cac453e8f3c41b812655c7e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=FSyhr-Ow7020W530ouX_9abfO8ry8Et-weh4zQ9BYVI";
let searchregardingapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/fb5e125f2eb640bf8aba86b15b9aeb03/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jlKc10_1MiM6CbGEbhF8xXITW2FVvbwzfAAF7ysaWzI";
let searchmissingemailsapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/37d8b1ec35454bfcbc5ca129c06823af/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=NFWR8XqT7yAhqtafw7rNmyYxqi6kLHwZlMHc3ybXNQ8";
let searchregardingopportunityapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7a6430c055e84319a1c69ae510f6bc0f/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-GhiiynMNUh0i_hKRXTGjSUsWJ6htoucfnBYN4rlLPQ";
  
async function loadControls() {
  loadMissingEmails();


  document.getElementById("searchBox").addEventListener("keydown", function(event) {
    if (event.key === "Enter") {
        event.preventDefault(); // Prevent default form submission
        const value = event.target.value;
       //showSuggestionsOnEnter(value); // Call the function
        
        filterOptions(event); // Call the function
    }
  });


  document.getElementById("searchBoxOpp").addEventListener("keydown", function(event) {
    if (event.key === "Enter") {
        event.preventDefault(); // Prevent default form submission
        const value = event.target.value;
        showOpportunitySuggestionsOnEnter(value); // Call the function
    }
  });
  Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {

    await getAttachmentsAsync(); 
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

document.querySelectorAll("#suggestions input[type='checkbox']")
    .forEach(cb => {
        cb.addEventListener("change", updateSelectedText);
    });



}

async function getAttachmentsAsync() {
  try {
    const mailattachments = Office.context.mailbox.item.attachments;
    console.log("Attachments: ");
    console.log(mailattachments);

    mailattachments.forEach(att => {
      console.log(att.id, att.name);
      Office.context.mailbox.item.getAttachmentContentAsync(
        att.id,
        { asyncContext: null },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const content = result.value.content;
            const format = result.value.format; // e.g., Base64
            console.log("Attachment content:", content);
            var attachment = {
              attachmentid: att.id,
              attachmentname: att.name,
              content: content,
              contentType: att.contentType,
              contentId: att.contentId,
              isInline: att.isInline,
            };
            attachments.push(attachment);
            //console.log("Attachments array: ", attachments);
            }
           else {
            console.error("Failed to get attachment content:", result.error.message);
          }
        }
      );
    });
  } catch (error) {
    console.error("Error retrieving attachments: ", error);
  }
}

function loadMissingEmails() {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      searchmissingemailsapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/37d8b1ec35454bfcbc5ca129c06823af/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=NFWR8XqT7yAhqtafw7rNmyYxqi6kLHwZlMHc3ybXNQ8";
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
      .catch((error) => console.error(error))
      .finally(() => {
          const searchmissingemailstitle = document.getElementById("searchmissingemailstitle");
          searchmissingemailstitle.style.display = "none";

          const trackbutton = document.getElementById("run");
          trackbutton.style.cursor = "pointer";
          trackbutton.style.pointerEvents = "auto";
          trackbutton.style.opacity = "1.0";
      });

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
        if (regardingItem == null) {
          regardingItem = {};
        }

        if (regardingItemOpp == null) {
          regardingItemOpp = {};
        }
        
        const inlineAttachments = attachments.filter(item => item.isInline === true);
        var description = result.value;

        inlineAttachments.forEach(item => {
          const contentId = item.contentId;
          const regex = new RegExp(`cid:${contentId}`, 'g');
          description = description.replace(regex, `data:${item.contentType};base64,${item.content}`);
        });

        const nonInlineAttachments = attachments.filter(item => item.isInline === false);
        // Successfully retrieved the email body
        const raw = JSON.stringify({
          from: from,
          to: to,
          cc: cc,
          subject: subject,
          description: description,
          useremailaddress: userEmailAddress,
          trackingid: trackingid,
          //dateTimeCreated: dateTimeCreated.format("YYYY-MM-DDTHH:mm:ss")
          dateTimeCreated: dateTimeCreatedUTC,
          regarding: selectedRegarding,
          regardingopp: regardingItemOpp,
          attachments: nonInlineAttachments
        });

        
  
        const requestOptions = {
          method: "POST",
          headers: myHeaders,
          body: raw,
          redirect: "follow",
        };

        createemailapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/4a3e289d2f3a48479a1fd674bbb3b5c1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=yzZ2_E4KcmDLpx_bzJ1AQtpC-H2mmuR0iyhGWCOJOxk";
  
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

function clearSuggestionsOpp(value) {
  const suggestionsDiv = document.getElementById("suggestionsopp");
  suggestionsDiv.innerHTML = "";
  suggestionsDiv.style.display = "none";
  if (value.length === 0) {
      // suggestionsDiv.style.display = "none";
      regardingItemOpp = null;
      console.log("regardingItemOpp: ", regardingItemOpp);
  }
}

function showSuggestionsOnEnter(value) {
  let suggestionsList = [];
  const suggestionsDiv = document.getElementById("suggestions");
  const searchText = document.getElementById("searchText");
  searchText.style.display = "block";

  searchregardingapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/fb5e125f2eb640bf8aba86b15b9aeb03/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jlKc10_1MiM6CbGEbhF8xXITW2FVvbwzfAAF7ysaWzI";

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


function showOpportunitySuggestionsOnEnter(value) {
  let suggestionsList = [];
  const suggestionsDiv = document.getElementById("suggestionsopp");
  const searchText = document.getElementById("searchTextOpp");
  searchText.style.display = "block";
  searchregardingopportunityapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7a6430c055e84319a1c69ae510f6bc0f/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-GhiiynMNUh0i_hKRXTGjSUsWJ6htoucfnBYN4rlLPQ"
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

fetch(searchregardingopportunityapi, requestOptions)
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
            div.classList.add("suggestion-itemopp");
            div.innerText = suggestion.name + " (" + suggestion.recordtype + ")";
            div.onclick = () => {
                document.getElementById("searchBoxOpp").value = suggestion.name + " (" + suggestion.recordtype + ")";
                regardingItemOpp = suggestion;
                suggestionsDiv.style.display = "none";
                console.log("regardingItemOpp: ", regardingItemOpp);
            };
            suggestionsDiv.appendChild(div);
        });
    } else {
        suggestionsDiv.style.display = "none";
    }

  })
  .catch((error) => console.error(error));


}


//START OF MULTISELECT DROPDOWN CODE

        const items = [

"Mid-Level Donor",
"Mid-Level Donor TT",
"Event Mid-Level Donor",
"Staff",
"Trustee",
"Volunteer",
"Corporate Partner",
"Sponsor",
"Major Donor",
"Foundation",
"Government",
"Board Member",
"Student",
"Alumni",
"Employee"

];

const selected = [];
var selectedRegarding = [];

var optionList = document.getElementById("optionList");
const selectedTags = document.getElementById("selectedTags");
const itemCount = document.getElementById("itemCount");

itemCount.innerHTML =  + "0 records";

function renderOptions(filter="",event){
  //toggleDropdown(event);
let suggestionsList = [];
   const searchText = document.getElementById("searchText");
  searchText.style.display = "block";

  searchregardingapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/fb5e125f2eb640bf8aba86b15b9aeb03/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jlKc10_1MiM6CbGEbhF8xXITW2FVvbwzfAAF7ysaWzI";

  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");

    const raw = JSON.stringify({
      "input": filter
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
      
      //suggestionsDiv.innerHTML = "";
      if (result.length === 0) {
          return;
      }
      else{
          suggestionsList.length = 0; // Clear the array
          result.forEach(item => {
              suggestionsList.push(item);
          });
      }

      itemCount.innerHTML = result.length + " records";
      
      // const filteredSuggestions = suggestionsList.filter(item => item.toLowerCase().startsWith(value.toLowerCase()));
      const filteredSuggestions = suggestionsList;
    

    optionList.innerHTML="";
    
    filteredSuggestions.sort().forEach(item=>{

        const label=document.createElement("label");
        const stringitem = JSON.stringify(item);
        label.className="option";

        label.innerHTML=`

        <div>

            <input type="checkbox"
                ${selected.includes(item.name)?"checked":""}
                onchange="toggleItem('${item.name}', '${item.recordid}', '${item.recordtype}')">

            ${item.name} (${item.recordtype})

        </div>

        `;

        optionList.appendChild(label);

    });

    toggleDropdown(event);

    // filteredSuggestions
    // .filter(x=>x.toLowerCase().includes(filter.toLowerCase()))
    // .forEach(item=>{

    //     const label=document.createElement("label");
    //     label.className="option";

    //     label.innerHTML=`

    //     <div>

    //         <input type="checkbox"
    //             ${selected.includes(item)?"checked":""}
    //             onchange="toggleItem('${item}')">

    //         ${item}

    //     </div>

    //     `;

    //     optionList.appendChild(label);

    // });

})}

//renderOptions();

function toggleDropdown(e){

    // e.stopPropagation();

    const d=document.getElementById("dropdown");

    d.style.display=
    d.style.display==="block"
    ?"none"
    :"block";
}

document.addEventListener("click",()=>{

document.getElementById("dropdown").style.display="none";

});

function toggleItem(name,recordid,recordtype){
   var item = { name: name, recordid: recordid, recordtype: recordtype };
    const index=selected.indexOf(name);

    if(index>-1)
    {
      selected.splice(index,1);
    }
    else{
      selected.push(item);
      selectedRegarding.push(item);
      console.log(selectedRegarding);
    }

    renderTags();
   // renderOptions(document.getElementById("searchBox").value);
}

function renderTags(){

    selectedTags.innerHTML="";

    selected.forEach(item=>{

        const div=document.createElement("div");

        div.className="tag";
        div.recordId = item.recordid;
        div.recordType = item.recordtype;

        div.innerHTML=`
        ${item.name}
        <span onclick="removeTag(event,'${item.name}', '${item.recordid}', '${item.recordtype}')">&times;</span>
        `;

        selectedTags.appendChild(div);

    });

}

function removeTag(e,item,recordid,recordtype){

  var selectedItem = { name: item, recordid: recordid, recordtype: recordtype };

    e.stopPropagation();

    const i = selected.findIndex(item => item.recordid === recordid);

    //const i=selected.indexOf(item);

    if(i>-1){

      selected.splice(i,1);
      selectedRegarding = selectedRegarding.filter(item => item.recordid !== recordid);
      console.log(selectedRegarding);
    }

    renderTags();
    renderOptions(document.getElementById("searchBox").value);

}

function filterOptions(){

    renderOptions(document.getElementById("searchBox").value,event);

}

document.getElementById("selectAll").addEventListener("change",function(){

    if(this.checked){

        selected.length=0;

        items.forEach(x=>selected.push(x));

    }else{

        selected.length=0;

    }

    renderTags();
    renderOptions(document.getElementById("searchBox").value);

});


function updateSelectedText() {
    const checked = [...document.querySelectorAll("#suggestions input:checked")]
        .map(cb => cb.value);

    document.getElementById("selectedText").textContent =
        checked.length
            ? checked.join(", ")
            : "Select accounts/contacts";
}