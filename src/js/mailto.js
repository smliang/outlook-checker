function registerProtocolHandler(){
    navigator.registerProtocolHandler("mailto",
                                  "https://outlook.office.com/?path=/mail/action/compose&to=%s",
                                  "Mailto: Links");
}

async function askRegisterProtocolHandler(){
    if(confirm("Do you want Outlook OWA to handle mailto: requests?")){
        //yes 
        console.log("registering protocol handler...");
        registerProtocolHandler();
        chrome.storage.local.set({protocolHandling: true});

    } else {
        chrome.storage.local.set({protocolHandling: false});
        chrome.contextMenus.create({
            title: "Turn on Mailto: Handling",
            id: "mailto on",
            contexts: ["action"]
        });
    }
}

chrome.storage.local.get(['protocolHandlingWait'], (data) => {
    console.log("protocol: ", data.protocolHandlingWait);
    if(data.protocolHandlingWait == true) {
        askRegisterProtocolHandler();
    }
    chrome.storage.local.set({protocolHandlingWait: false});
});