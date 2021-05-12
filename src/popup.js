var cache = {};
const outlookUrl = 'https://outlook.office.com/mail';
const outlookHost = 'outlook.office.com';

document.addEventListener('DOMContentLoaded', changeName());

function changeName() {
    chrome.storage.local.get(['token', 'displayName', 'email', 'unreadCount'], (res) => {
        console.log('got info from cache: ', res);
        cache.token = res.token;
        cache.displayName = res.displayName;
        cache.email = res.email;
        cache.unread = res.unreadCount;

        //change info on top bar
        document.getElementById("user").innerHTML = cache.displayName;
        document.getElementById("email").innerHTML = cache.email;
        document.getElementById("numUnread").innerHTML = cache.unread;
        
        let links = document.getElementsByClassName("go-to-inbox");
        console.log("links!: ", links);
        for(link of links){
            link.addEventListener('click', openOutlookInbox);
        }
        getEmails();
    })
}

async function getEmails(){
    let res = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$filter=isRead eq false', {
        headers: new Headers({
            'Authorization': 'Bearer ' + cache.token,
            'Content-Type': 'application/json'
        })
    });

    if (!res.ok) {
        await console.log(res.json());
        throw new Error('HTTP error: status = ' + res.status);
    }

    res = await res.json();
    console.log("emails: ",res.value);
    cache.allEmails = res.value;
    updateEmailDisplay();

}
 
function updateEmailDisplay(){
    var emailHTML = "";
    for(email of cache.allEmails){
        //unread email, post!
        if(!email.isRead){
            let initial = email.from.emailAddress.name.match("[a-zA-Z]");
            console.log("initial: " + initial);

            let html = `<div class="row email">`+
                            `<div class="col-1 icon">`+
                                `<h1 class="initial">${initial}</h1>`+
                            `</div>`+
                            `<div class="col">`+
                                `<p class="sender">${email.from.emailAddress.name}<span style="font-style:italic; font-weight:normal"> ${email.from.emailAddress.address}</span></p>`+
                                `<p class="subject">${email.subject}</p>`+
                                `<p class="body">${email.bodyPreview}...</p>`+
                            `</div>`+
                        `</div>`
            emailHTML = emailHTML + html;
         //   console.log("EMAIL: ", html);
        }
        //console.log("string of html: " + emailHTML);

        //append to innerHTML of id="emails"
        document.getElementById("emails").innerHTML = emailHTML;
        document.getElementById("numUnread").innerHTML = cache.allEmails.length;
    }
}

async function openOutlookInbox() {
    let windowIdToFocus = await new Promise((resolve) => {
        chrome.tabs.query({
            url: `*://${outlookHost}/*`
        }, tabs => {
            if (tabs.length > 0) {
                chrome.tabs.update(tabs[tabs.length - 1].id, {active: true});
                resolve(tabs[tabs.length - 1].windowId);
            } else {
                chrome.tabs.create({url: outlookUrl});
                resolve();
            }
        });
    });
    if(windowIdToFocus) {
        await new Promise((resolve) => {
            chrome.windows.update(windowIdToFocus, {
                focused: true
            }, () => {
                resolve();
            });
        });
    }
}
