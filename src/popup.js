var cache = {};
const outlookUrl = 'https://outlook.office.com/mail';
const outlookHost = 'outlook.office.com';

document.addEventListener('DOMContentLoaded', changeName());
//todo add an event handler to update popup in case its open when we need a token refresh


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

        if(cache.unread){
            let el =  document.getElementById("numUnread");
            el.innerHTML = cache.unread;
            el.parentElement.style.visibility = "visible";
        }
        
        
        let links = document.getElementsByClassName("go-to-inbox");
        console.log("links!: ", links);
        for(link of links){
            link.addEventListener('click', openOutlookInbox);
        }
        getEmails();
    })
}


//gets 25 emails
async function getEmails(){

    let res = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=25&$filter=isRead eq false', {
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
    var i = 0;
    for( email of cache.allEmails){
        //unread email, post!
        if(!email.isRead){
            //console.log("email: ", i);
            let initial = email.from.emailAddress.name.match("[a-zA-Z]");
            let sendDateTime = new Date(email.sentDateTime); // "2021-05-11T18:24:08Z"
            let minute = sendDateTime.getMinutes().toString();
            let d = sendDateTime.getHours()%12 + ":" + (minute < 10 ? "0"+ minute : minute);
            d = d + (sendDateTime.getHours() > 12 ? " pm" : " am");
            
            //for setting up flagged status
            let flagged = (email.flag.flagStatus == "flagged" ? true : false);
            let iconClass = (flagged ? "ms-Icon ms-Icon--EndPointSolid flagged" : "ms-Icon ms-Icon--Flag unflagged");
            let toolTip = (flagged ? "Unflag" : "Flag");
            let subStyle = (flagged ? "subject blue-bold" : "subject")
            

            let html = `<div class="row email" data-idx=${i}>`+
                            `<div class="padding">
                                <div class="col-1 icon">`+
                                    `<h1 class="initial">${initial}</h1>`+
                                `</div>`+
                                `<div class="col">`+
                                    `<div class="row">`+
                                            `<div class="col sender-box">`+
                                                `<p class="sender">${email.from.emailAddress.name}<span style="font-style:italic; font-weight:normal"> ${email.from.emailAddress.address}</span></p>`+
                                                `<p class="${subStyle}">${email.subject}</p>`+
                                            `</div>`+
                                            `<div class="col-4 mail-icons mt-1">
                                                <button class="btn icon-button readToggle" title="Mark as Read"><i class="ms-Icon ms-Icon--Read unread"></i></button>
                                                <button class="btn icon-button archive" title="Archive"><i class="ms-Icon ms-Icon--Archive"></i></button>    
                                                <button class="btn icon-button delete" title="Delete"><i class="ms-Icon ms-Icon--Delete"></i></button>    
                                                <button class="btn icon-button flag" title="${toolTip}"><i class="${iconClass}"></i></button>    
                                            </div>`+
                                            `<div class="col-2 time-box">`+
                                                `<p class="time">${d}</p>`+
                                            `</div>`+
                                    `</div>`+
                                    
                                    
                                    `<p class="body">${email.bodyPreview}...</p>`+
                                    
                                `</div>`+
                            `</div>`+
                        `</div>`
            emailHTML = emailHTML + html;
            //console.log("EMAIL: ", html);
            ++i;
        }
        //console.log("string of html: " + emailHTML);

        //append to innerHTML of id="emails"
        document.getElementById("emails").innerHTML = emailHTML;

        //init toolbar
            //unread/read toggle
        let readToggles = document.getElementsByClassName("readToggle")
        for(readToggle of readToggles){
            readToggle.addEventListener('click', toggleRead);
        }
            //archive button
        let archiveButtons = document.getElementsByClassName("archive");
        for(button of archiveButtons){
            button.addEventListener('click', archiveMessage);
        }
            //delete button
        let deleteButtons = document.getElementsByClassName("delete");
        for(button of deleteButtons){
            button.addEventListener('click', deleteMessage);
        }
            //flag/unflag toggle
        let flagButtons = document.getElementsByClassName("flag");
        for(button of flagButtons){
            button.addEventListener('click', toggleFlag);
        }

        //update counts because we only get 25 emails, so this will be more accurate unless there's more than 25 unread
        if(cache.unread < 25 ){
            setUnreadCount(cache.allEmails.length);
            cache.unread = cache.allEmails.length;
        }
    }
}

//changes unread badge and icon badge
function setUnreadCount(count) {
    let el =  document.getElementById("numUnread");
    if(count){
        el.innerHTML = count;
        el.parentElement.style.visibility = "visible";
    }
    else el.parentElement.style.visibility = "collapse";

    chrome.action.setBadgeBackgroundColor({ color: [208, 0, 24, 255] });
    console.log("count: "+count);
    chrome.action.setBadgeText({
        text: count === 0 ? '' : count.toString()
    });
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

function toggleRead(e){
    var icon = e.target.firstChild;
    var idx = e.path[5].dataset.idx;
    var email = cache.allEmails[idx];
    var sender = e.path[2].firstChild.firstChild;
    var subj = e.path[2].firstChild.lastChild;
    //console.log("subj: ", subj);


    console.log("unread clicked!", e, icon,idx,email);

    toggleReadIcon(icon, sender, subj);
    sendReadUpdate(idx);
    
}

function toggleReadIcon(iconref, sendref, subjref){
    //check if in unread or read state
    var classType = iconref.className;
    if(classType.search("unread") != -1){
        //change to open envelope and read
        iconref.className = "ms-Icon ms-Icon--Mail read"
        iconref.title = "Mark as Unread";
        sendref.style.fontWeight = "normal";
        subjref.style.fontWeight = "normal";
        setUnreadCount(--cache.unread);
    }
    else{
        iconref.className = "ms-Icon ms-Icon--Read unread"
        iconref.title = "Mark as Read";
        sendref.style.fontWeight = "600";
        subjref.style.fontWeight = "600";
        setUnreadCount(++cache.unread);
    }

}

async function sendReadUpdate(idx){
    var newReadValue = !cache.allEmails[idx].isRead;
    var id = cache.allEmails[idx].id;
    console.log("email read is now ", newReadValue);
    
    let res = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/' + id, {
        method: 'PATCH',
        body: JSON.stringify({
            isRead: newReadValue
        }),
        headers: new Headers({

            'Authorization': 'Bearer ' + cache.token,
            'Content-Type': 'application/json'
        })
    });

    if (!res.ok) {
        await console.log(res.json());
        throw new Error('HTTP error: status = ' + res.status);
    }

    cache.allEmails[idx].isRead = !cache.allEmails[idx].isRead;

    console.log("finished talking to update read");
    res = await res.json();
    console.log(res);
}

function archiveMessage(e){
    console.log("archive clicked");
    var idx = e.path[5].dataset.idx;
    var emailDiv= e.path[5];

    console.log(emailDiv);
    emailDiv.style.maxHeight = 0;

    console.log("archiving idx: ", idx);
    sendArchiveUpdate(idx);
}

async function sendArchiveUpdate(idx){
    var id = cache.allEmails[idx].id;
    let res = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/' + id +"/move", {
        method: 'POST',
        body: JSON.stringify({
            destinationId: "archive"
        }),
        headers: new Headers({

            'Authorization': 'Bearer ' + cache.token,
            'Content-Type': 'application/json'
        })
    });

    if (!res.ok) {
        await console.log(res.json());
        throw new Error('HTTP error: status = ' + res.status);
    }

    console.log("send message to archive", res.json());
    cache.allEmails[idx].parentFolderId = cache.archive;
    setUnreadCount(--cache.unread);
}


function deleteMessage(e){
    console.log("delete clicked");
    var idx = e.path[5].dataset.idx;
    var emailDiv= e.path[5];

    console.log(emailDiv);
    emailDiv.style.maxHeight = 0;

    console.log("delete idx: ", idx);
    sendDeleteUpdate(idx);
}

async function sendDeleteUpdate(idx){
    var id = cache.allEmails[idx].id;
    let res = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/' + id , {
        method: 'DELETE',
        headers: new Headers({

            'Authorization': 'Bearer ' + cache.token,
            'Content-Type': 'application/json'
        })
    });

    if (!res.ok) {
        await console.log(res.json());
        throw new Error('HTTP error: status = ' + res.status);
    }

    console.log("send deleded message successful");
    cache.allEmails[idx].parentFolderId = cache.archive;
    setUnreadCount(--cache.unread);
}


function toggleFlag(e){
    var icon = e.target.firstChild;
    var idx = e.path[5].dataset.idx;
    var sender = e.path[2].firstChild.firstChild;
    var subj = e.path[2].firstChild.lastChild;
    //console.log("subj: ", subj);


    //console.log("flag clicked!", e, icon,idx,email);

    toggleFlagIcon(icon, subj);
    sendFlagUpdate(idx);
    
}

function toggleFlagIcon(iconref, subjref){
    console.log("toggle flag state", iconref);
    var classType = iconref.className;
    if(classType.search("unflagged") != -1){ //unflagged, change to flagged
        console.log("toggle unflag to flag")
        iconref.className = "ms-Icon ms-Icon--EndPointSolid flagged";
        iconref.title = "Unflag";
        subjref.className = "subject blue-bold";
        
    }
    else{ 
        //change to flagged: change to unflag
        console.log("toggle flag to unflagged: ", iconref.className);
        iconref.className = "ms-Icon ms-Icon--Flag unflagged";
        iconref.title = "Flag";
        subjref.className = "subject";
        
    }

}

async function sendFlagUpdate(idx){
    var newFlagValue = (cache.allEmails[idx].flag.flagStatus == "notFlagged" ? "flagged" : "notFlagged");
    var id = cache.allEmails[idx].id;
    console.log("email flag is now ", newFlagValue);
    
    let res = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/' + id, {
        method: 'PATCH',
        body: JSON.stringify({
            flag: {
                flagStatus: newFlagValue
            }
        }),
        headers: new Headers({

            'Authorization': 'Bearer ' + cache.token,
            'Content-Type': 'application/json'
        })
    });

    if (!res.ok) {
        await console.log(res.json());
        throw new Error('HTTP error: status = ' + res.status);
    }

    cache.allEmails[idx].flag.flagStatus = newFlagValue;

    console.log("finished talking to update flag");
    res = await res.json();
    console.log(res);
}