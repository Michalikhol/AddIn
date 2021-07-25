/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/** @type {Event} */
var sendEvent;
/** @type {string[]} */
var groups = [];
/** @type {Office.Message} */
var mailboxItem;
/** @type {string} */
var thisUser= "";
/** @type {boolean} */
var isExternal = false;
/** @type {Office.MessageCompose} */
var composeItem;

var dialog;


    Office.initialize = function (reason) {
        mailboxItem = Office.context.mailbox.item;
        composeItem = Office.cast.item.toItemCompose(mailboxItem);
    }

    function checkRecipients(event) {
      sendEvent = event;
      differntClassifiction(event)
    }

    function differntClassifiction(event) {
      composeItem.to.getAsync(
        {asyncContext: event},
        function (asyncResult) {

          var recipients = asyncResult.value

          // if it find user that matches the regexp it exits with true
          isExternal = checkClassification(recipients)

          if(isExternal){
            getResponseFromUser()
          } else {

            composeItem.cc.getAsync(
              {asyncContext: event},
              function (asyncResult) {
                var ccRecipients = asyncResult.value
                  if(ccRecipients.length > 0) {
                      isExternal = checkClassification(ccRecipients)

                      if(isExternal) {
                        getResponseFromUser()
                      } else {
                        asyncResult.asyncContext.completed({ allowEvent: true });
                      }
                  } else {
                    asyncResult.asyncContext.completed({ allowEvent: true });
                  }
              }
            )
          }
        }
      )
    }

    function checkClassification(recipientsToCheck) {
      var found = false
      found =  recipientsToCheck.some(recipient => {
        var checkTo = (new RegExp(/[a-z]*@[a-z]*.meimad/gm).test(recipient.emailAddress))
        if(checkTo) {
          return true
        }
      })

      return found
    }

   
    function getResponseFromUser() {
      var url = `${window.location.origin}/dialog.html`;
      var dialogOptions = { width: 30, height: 20 ,  displayInIframe: true };
  
      Office.context.ui.displayDialogAsync(url, dialogOptions,
         function(asyncResult) {
           if(asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            dialog= asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
            dialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
           } else {
             console.log("hi")
            dialogClosed()
           }
          
      });
    }
    
      function receiveMessage(arg) {
        var messageFromDialog = JSON.parse(arg.message);

        if(messageFromDialog) {
          dialog.close();
          dialog = null;
          removeProgress();
          // SEND
          sendEvent.completed({ allowEvent: true });
        } 
        else {
          dialog.close();
          dialog = null;
          removeProgress();
          sendEvent.completed({ allowEvent: false });
        }
      }
      
    function dialogClosed() {
      dialog = null;
      removeProgress();
      sendEvent.completed({ allowEvent: false });
    }
  
    function removeProgress() {
      mailboxItem.notificationMessages.removeAsync("progress");
    }

  