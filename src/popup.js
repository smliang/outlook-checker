var cache = {};
const outlookUrl = 'https://outlook.office.com/mail';
const outlookHost = 'outlook.office.com';

document.addEventListener('DOMContentLoaded', changeName());
//init big buttons
let refreshButton = document.getElementById("refresh-button");
refreshButton.addEventListener('click', onRefresh);

let readButton = document.getElementById("mark-all");
readButton.addEventListener('click', markAllAsRead);
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
            email.displayed = true;
            let initial = email.from.emailAddress.name.match("[a-zA-Z]");
            let color = getRandomColor();
            email.color = color;
            let sendDateTime = new Date(email.sentDateTime); // "2021-05-11T18:24:08Z"
            let minute = sendDateTime.getMinutes().toString();
            let hour = sendDateTime.getHours()%12;
            let d = (hour == 0 ? 1 : hour) + ":" + (minute < 10 ? "0"+ minute : minute);
            d = d + (sendDateTime.getHours() > 12 ? " pm" : " am");
            
            //for setting up flagged status
            let flagged = (email.flag.flagStatus == "flagged" ? true : false);
            let iconClass = (flagged ? "ms-Icon ms-Icon--EndPointSolid flagged" : "ms-Icon ms-Icon--Flag unflagged");
            let toolTip = (flagged ? "Unflag" : "Flag");
            let subStyle = (flagged ? "subject blue-bold" : "subject")
            let bgColor = (flagged ? "style='background-color:lightyellow'" : "")
            

            let html = `<div class="row email" data-idx=${i} ${bgColor}>`+
                            `<div class="col unread-bar"></div>` +
                            `<div class="col padding">`+
                                `<div class="col-1 icon">`+
                                    `<h1 class="initial ms-bgColor-shared${color}">${initial}</h1>`+
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
        } else { email.displayed = false; }
        //console.log("string of html: " + emailHTML);

        //append to innerHTML of id="emails"
        let emailsDiv = document.getElementById("emails");
        emailsDiv.innerHTML = emailHTML;
        setTimeout( () =>{ emailsDiv.className = "col expand"}, 100) ;

        //init email toolbars
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

        //email click
        emailsDiv.addEventListener('click', expandEmailView);

        //update counts because we only get 25 emails, so this will be more accurate unless there's more than 25 unread
        if(cache.unread < 25 ){
            setUnreadCount(cache.allEmails.length);
            cache.unread = cache.allEmails.length;
        }
    }
}

function onRefresh(e){
    console.log("refreshing email display!",e);
    getEmails();
}

//changes unread badge and icon badge
function setUnreadCount(count) {
    let el =  document.getElementById("numUnread");
    if(count > 0){
        el.innerHTML = count;
        el.parentElement.style.visibility = "visible";
    }
    else el.parentElement.style.visibility = "collapse";

    chrome.action.setBadgeBackgroundColor({ color: [208, 0, 24, 255] });
    console.log("display unread count: "+count);
    chrome.action.setBadgeText({
        text: count === 0 ? '' : count.toString()
    });
}

function expandEmailView(e){    
    if(e.target.className.search("icon-button") == -1){
        console.log("expanding email!", e);
        openEmailPreview(e.target.dataset.idx);
    }
}

async function openEmailPreview(idx){
    //init template
    initExpandedTemplate(idx);
    console.log("animating")

    //so i'd really like for this to swipe over, but i simply cannot get the animations how i want them
    //so we're just gonna snap and maybe I'll come back to it later... see expand-view for the animations i 
    //commented out
    let emailDiv = document.getElementById("emails");
    emailDiv.className = "col expand expand-view";
    document.getElementById("toolbar-buttons").style.display = "none";

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
    var bar = e.path[5].firstChild;
    //console.log("subj: ", subj);


    console.log("unread clicked!", e, icon,idx,email);

    toggleReadIcon(icon, sender, subj, bar);
    sendReadUpdate(idx);
    
}

function toggleReadIcon(iconref, sendref, subjref, barref){
    //check if in unread or read state
    var classType = iconref.className;
    if(classType.search("unread") != -1){
        //change to open envelope and read
        iconref.className = "ms-Icon ms-Icon--Mail read"
        iconref.title = "Mark as Unread";
        sendref.style.fontWeight = "normal";
        subjref.style.color = "black";
        barref.className = "unread-bar read"
        setUnreadCount(--cache.unread);
    }
    else{
        iconref.className = "ms-Icon ms-Icon--Read unread"
        iconref.title = "Mark as Read";
        sendref.style.fontWeight = "600";
        subjref.style.color = "#0078d4";
        barref.className = "unread-bar"
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

async function markAllAsRead(){
    console.log("marking all as read");
    let emails = document.getElementsByClassName("email");
    //console.log("emails: ", emails);
    for(emailDiv of emails){
        let i = emailDiv.dataset.idx
        let email = cache.allEmails[i];
        if(email.displayed > 0 && !email.isRead){ //grab emails that havent been read, archived, or deleted
            
            let icon = emailDiv.getElementsByClassName("ms-Icon")[0];
            let subj = emailDiv.getElementsByClassName("subject")[0];
            let sender = emailDiv.getElementsByClassName("sender")[0];
            let bar = emailDiv.firstChild;
            console.log("icon: ", icon, subj, sender);

            toggleReadIcon(icon, sender, subj, bar);
            sendReadUpdate(i);
        }
    }
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
    cache.allEmails[idx].displayed = false;
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
    cache.allEmails[idx].displayed = false;
    setUnreadCount(--cache.unread);
}


function toggleFlag(e){
    var icon = e.target.firstChild;
    var idx = e.path[5].dataset.idx;
    var emailDiv = e.path[5];
    var subj = e.path[2].firstChild.lastChild;
    //console.log("subj: ", subj);


    //console.log("flag clicked!", e, icon,idx,email);

    toggleFlagIcon(icon, subj, emailDiv);
    sendFlagUpdate(idx);
    
}

function toggleFlagIcon(iconref, subjref, emailref){
    console.log("toggle flag state", iconref);
    var classType = iconref.className;
    if(classType.search("unflagged") != -1){ //unflagged, change to flagged
        console.log("toggle unflag to flag")
        iconref.className = "ms-Icon ms-Icon--EndPointSolid flagged";
        iconref.title = "Unflag";
        subjref.className = "subject blue-bold";
        emailref.style.backgroundColor = "lightyellow"
        
    }
    else{ 
        //change to flagged: change to unflag
        console.log("toggle flag to unflagged: ", iconref.className);
        iconref.className = "ms-Icon ms-Icon--Flag unflagged";
        iconref.title = "Flag";
        subjref.className = "subject";
        emailref.style.backgroundColor = "white"
        
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

const colors = [
    "pinkRed10",
    "red20",
    "red10",
    "orange20",
    "orangeYellow20",
    "green10",
    "green20",
    "cyan20",
    "cyan30",
    "cyanBlue20",
    "blue10",
    "blueMagenta30",
    "blueMagenta20",
    "magenta20",
    "magenta10",
    "magentaPink20",
    "orange30",
    "gray30",
    "gray20"
]
//set random color of icons
function getRandomColor(){
    let len = colors.length;
    return(colors[Math.floor(Math.random() * len)]);
}

async function initExpandedTemplate(idx){
        let email = cache.allEmails[idx];
        //console.log(template);

        //flag status
        if(email.flag.flagStatus == "flagged"){
            document.getElementById("sender-box-expanded").style.backgroundColor="lightyellow";
            //TODO: set flag icon
        }

        // subject 
        document.getElementById("subject-expanded").innerHTML = email.subject;
        // set up icon
        let icon = document.getElementById('expanded-initial');
        icon.class = `initial ms-bgColor-shared${email.color}`;
        icon.innerHTML = email.from.emailAddress.name.match("[a-zA-Z]");

        //from
        document.getElementById("sender-expanded").innerHTML = `${email.from.emailAddress.name} &lt${email.from.emailAddress.address}&gt`

        //time
        document.getElementById("time-expanded").innerHTML = getParsedTime(email.sentDateTime);

        //to
        let to =  document.getElementById("to-expanded");
        let recipients = `<span style="font-weight:600">To:</span>`;
        for(recipient of email.toRecipients){ 
            let email = recipient.emailAddress.address;
            let name = recipient.emailAddress.name;
            let line = ` ${name} <span class="font-italic">&lt${email}&gt</span>,`;
            recipients = recipients + line;
            console.log("line: ", line);
        }
        
        recipients = recipients.substring(0, recipients.length-1);
        console.log("recipients:", recipients);
        to.innerHTML = recipients;

        document.getElementById("body-expanded").innerHTML = `${email.body.content}`;
        console.log("finished template init");
        
    
    
}

const days = ["Sun", "Mon", "Tue", "Wed", "Thur", "Fri", "Sat"];
function getParsedTime(dateObj){
    let sendDateTime = new Date(dateObj); // "2021-05-11T18:24:08Z"
    let day = days[sendDateTime.getDay()];
    
    let month = sendDateTime.getMonth() + 1;
    let date = sendDateTime.getDate();
    let year = sendDateTime.getFullYear();

    let minute = sendDateTime.getMinutes().toString();
    let hour = sendDateTime.getHours()%12;
    let time = (hour == 0 ? 1 : hour) + ":" + (minute < 10 ? "0"+ minute : minute);
    time = time + (sendDateTime.getHours() > 12 ? "pm" : "am");

    return `${day} ${month}/${date}/${year} ${time}`;
}