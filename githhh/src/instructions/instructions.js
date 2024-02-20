// Office.onReady(function (info) {
//     // Register to receive messages from the parent window
//     Office.context.messageParent(function (message) {
//         var receivedData = message.message;
//         // Handle the received data as needed
//         console.log(receivedData);
//     });

//     // Now the child window is ready to receive messages from the parent
// });
Office.onReady(function (info) {
    console.log("mani");
    if (info.isOfficeInitialized) {
        console.log("mani");
      // Get the data object sent from the parent
      OfficeExtension.ExtensionHelpers.getDialogMessage().then(function (messageFromParent) {
        if (messageFromParent) {
          // Use messageFromParent as your object
          console.log(messageFromParent);
        }
      });
    }
  });
  