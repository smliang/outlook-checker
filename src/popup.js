//applicaiton ID 7041e3fb-49c9-454c-86f1-ce7bcaee8db4
//client secret value 21vn7M~MA_NieB3lm6Toon.f19VuQp~C45

chrome.runtime.onMessage.addListener(onMessage);

function onMessage(req, sender, reply){
    if(!req.login){
        window.close();
        console.log("logout!");
    }
}