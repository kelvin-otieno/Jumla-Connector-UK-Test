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


function loadMissingEmails() {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      searchmissingemailsapi = "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/04/workflows/2fea241966174c12944f9a3b88ac1a0c/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=gUZ6iTRC3iqxdlbG6DEaEoCucGwwk496JWLgiwicmQE";
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
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    const from = item.from;
    const to = item.to;
    const cc = item.cc;
    const subject = item.subject;
    const trackingid = item.conversationId;
    const dateTimeCreated = item.dateTimeCreated;
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
          dateTimeCreated: dateTimeCreated
        });
  
        const requestOptions = {
          method: "POST",
          headers: myHeaders,
          body: raw,
          redirect: "follow",
        };
  
        fetch(
          "https://a26068ef5a2445e0ad4ddab310c157.f9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a1a72d0bb7d9448f8da9f95bfa768fd6/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=SEivYIWI91pJ1IWu8Jr2T-GiziUTSUNkSi-7IY0wGtk",
          requestOptions
        )
          .then((response) => {
            if (response.ok) {
              const label = document.createElement("b");

              const link = document.createElement("a");
              link.href = "https://example.com/emails/123"; // where "Email" should point
              link.textContent = "Email";

              label.appendChild(link);
              label.appendChild(document.createTextNode(" successfully created."));
              // let label = document.createElement("b").appendChild(document.createTextNode("Email successfully created."));
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
  