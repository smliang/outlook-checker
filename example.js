//FRON not_logged_in_popup.js
function login() {
    window.chrome.runtime.sendMessage({ login: true });
    window.close();
}
document.getElementById("login_button").addEventListener('click', login);


//FROM background.js
const refreshWindowMs = 10 * 60 * 1000;

const outlookUrl = 'https://outlook.office.com/mail';
const outlookHost = 'outlook.office.com';
const adClientId = 'b766c6fc-a3a3-4f17-a0d5-650072398ef0';
const tokenEndpoint = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
const outlookHostFilter = {
    url: [{hostEquals: outlookHost}]
};
let mainMenuId;

initListeners();
if (localStorage.getItem('token') === null) {
    setUiToLogoutState().then(async () => {
        await update();
    });
} else {
    setUiToLoggedInState().then(async () => {
        await update();
    });
}

function initListeners() {
    chrome.runtime.onInstalled.addListener(onInstalled);
    chrome.alarms.onAlarm.addListener(onAlarm);
    chrome.webNavigation.onHistoryStateUpdated.addListener(onNavigate, outlookHostFilter);
    chrome.browserAction.onClicked.addListener(onClicked);
    chrome.runtime.onMessage.addListener(onMessage);
    chrome.contextMenus.onClicked.addListener(onContextMenuClicked);
}

function onInstalled() {
    chrome.alarms.create('update', {periodInMinutes: 1});
}

async function onAlarm(alarm) {
    await update();
}

async function onNavigate(details) {
    if (details.url && isOutlookUrl(details.url)) {
        await update();
    }
}

async function onClicked() {
    if (localStorage.getItem('token') !== null) {
        await openOutlook();
    }
}

async function onMessage(request, sender, sendResponse) {
    if (request.login) {
        await login();
        await update();
    }
}

async function onContextMenuClicked(info) {
    if (info.menuItemId === "MainMenu") {
        await logOut();
    }
}

async function update() {
    console.log('update called');
    await refreshTokenIfNeeded();
    const count = await getInboxCount();
    setUnreadCount(count);
}

async function refreshTokenIfNeeded() {
    const token = localStorage.getItem('token');
    if (token === null) {
        throw new Error('Unable to refresh token. Not logged in.');
    }
    const expiration = new Date(parseInt(localStorage.getItem('expiration')));
    let refreshTime = new Date(expiration.getTime() - refreshWindowMs);
    if (new Date() >= refreshTime) {
        await refreshToken();
    }
}

async function refreshToken() {
    console.log('refreshToken called');
    console.log(`token: ${localStorage.getItem('refresh_token')}`);
    let res = await fetch(tokenEndpoint, {
        method: 'POST',
        body: new URLSearchParams({
            'client_id': adClientId,
            'scope': 'Mail.Read offline_access',
            'refresh_token': localStorage.getItem('refresh_token'),
            'grant_type': 'refresh_token'
        })
    });
    if (!res.ok) {
        await logOut();
        throw new Error('HTTP error, status = ' + res.status);
    }
    res = await res.json();
    saveToken(res);
}

async function getInboxCount() {
    console.log('getInboxCount called');
    let res = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox?$select=unreadItemCount', {
        headers: new Headers({
            'Authorization': 'Bearer ' + localStorage.getItem('token'),
            'Content-Type': 'application/json'
        })
    });
    if (!res.ok) {
        console.log('getInboxCount res.ok != true');
        if (res.status === 401) {
            console.log('getInboxCount res.status = 401');
            await logOut();
        }
        throw new Error('HTTP error, status = ' + res.status);
    }
    res = await res.json();
    return res.unreadItemCount;
}

function setUnreadCount(count) {
    chrome.browserAction.setBadgeBackgroundColor({color: [208, 0, 24, 255]});
    chrome.browserAction.setBadgeText({
        text: count === 0 ? '' : count.toString()
    });
}

function isOutlookUrl(url) {
    return url.indexOf(outlookUrl) === 0;
}

async function login() {
    const state = (Date.now() * Math.random()).toString();
    const redirectUrl = chrome.identity.getRedirectURL();
    console.log(redirectUrl);
    const codeVerifier = createRandomString();
    const hash = await sha256(codeVerifier);
    const codeChallenge = bufferToBase64UrlEncoded(hash);
    var launchWebAuthFlow = new Promise((resolve) => {
        chrome.identity.launchWebAuthFlow({
            url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?' +
                'response_type=code' +
                '&response_mode=query' +
                `&client_id=${adClientId}` +
                `&redirect_uri=${redirectUrl}` +
                '&prompt=select_account' +
                '&scope=Mail.Read offline_access' +
                `&state=${state}` +
                '&code_challenge_method=S256' +
                `&code_challenge=${codeChallenge}`,
            interactive: true
        }, (redirectUrl) => {
            resolve(redirectUrl);
        });
    });
    const responseUrl = await launchWebAuthFlow;
    var params = (new URL(responseUrl)).searchParams;
    if (params.get("state") !== state)
        return;
    const code = params.get("code");
    let res = await fetch(tokenEndpoint, {
        method: 'POST',
        body: new URLSearchParams({
            'client_id': adClientId,
            'scope': 'Mail.Read offline_access',
            'code': code,
            'redirect_uri': redirectUrl,
            'grant_type': 'authorization_code',
            'code_verifier': codeVerifier
        })
    });
    if (!res.ok) {
        throw new Error('HTTP error, status = ' + res.status);
    }
    res = await res.json();
    saveToken(res);
    await setUiToLoggedInState();
}

function clearStorage() {
    localStorage.removeItem('token');
    localStorage.removeItem('refresh_token');
    localStorage.removeItem('expiration');
    localStorage.removeItem('unreadCount');
}

async function logOut() {
    console.log('logOut called');
    clearStorage();
    await setUiToLogoutState();
}

async function openOutlook() {
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
    await update();
}

function saveToken(res) {
    console.log('saveToken called');
    console.log(`token ${res.access_token}`);
    console.log(`token ${res.refresh_token}`);
    localStorage.setItem('token', res.access_token);
    localStorage.setItem('refresh_token', res.refresh_token);
    const expiration = new Date(Date.now() + parseInt(res.expires_in) * 1000);
    console.log(`expiration ${expiration} (expiration.getTime().toString())`);
    localStorage.setItem('expiration', expiration.getTime().toString());
}

async function setUiToLoggedInState() {
    chrome.browserAction.setIcon({
        path: {
            "16": "images/icon_16px_logged_in.png",
            "32": "images/icon_32px_logged_in.png",
            "48": "images/icon_48px_logged_in.png",
            "128": "images/icon_128px_logged_in.png"
        }
    });
    chrome.browserAction.setBadgeBackgroundColor({color: [208, 0, 24, 255]});
    chrome.browserAction.setBadgeText({text: ''});
    chrome.browserAction.setPopup({'popup': ''});
    await showMenu();
}

async function setUiToLogoutState() {
    chrome.browserAction.setIcon({path: 'images/notLoggedInIcon.png'});
    chrome.browserAction.setBadgeBackgroundColor({color: [190, 190, 190, 230]});
    chrome.browserAction.setBadgeText({text: ''});
    chrome.browserAction.setPopup({'popup': 'popup/html/NotLoggedInPopup.html'});
    await hideMenu();
}

async function showMenu() {
    await hideMenu();
    mainMenuId = chrome.contextMenus.create({
        id: "MainMenu",
        title: "Sign out",
        contexts: ['browser_action']
    });
}

function hideMenu() {
    return new Promise((resolve) => {
        chrome.contextMenus.removeAll(function () {
            resolve();
        })
    });
}
