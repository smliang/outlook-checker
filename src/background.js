const refreshWindowMS = 10 * 60 * 1000;

const outlookUrl = 'https://outlook.office.com/mail';
const outlookHost = 'outlook.office.com';
const adClientId = '7041e3fb-49c9-454c-86f1-ce7bcaee8db4';
const tokenEndpoint = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
const outlookHostFilter = {
    url: [{ hostEquals: outlookHost }]
};
let mainMenuId;


chrome.runtime.onInstalled.addListener(function () {
    //set to logout state
    console.log("hello");

    setUILogout();
    chrome.alarms.create('update', { periodInMinutes: .5 });

});

chrome.runtime.onMessage.addListener(onMessage);
chrome.alarms.onAlarm.addListener(onAlarm);

function setUILogout() {
    chrome.storage.local.set({ login: false });
    //chrome.storage.local.get(['login'], function(result){console.log(result.login)});
    //TODO: add new loggedout icon chrome.browserAction.setIcon({path: 'PUTPATHHERE'});
    // chrome.browserAction.setBadgeText({text: ''});

    //change to login popup
    chrome.action.setPopup({ 'popup': 'popup_logged_out.html' });
}

function setUILogin() {
    chrome.storage.local.set({ 'login': true });
    //chrome.action.setIcon({path: 'icons/icon.png'}); //TODO: add more icon sizes
    //   chrome.browserAction.setBadgeBackgroundColor({color: [208, 0, 24, 255]});
    //   chrome.browserAction.setBadgeText({text: ''});
    chrome.action.setPopup({ 'popup': 'popup.html' });
    update();
}

//handler to log in user
async function onMessage(request, sender, sendResponse) {
    //console.log("got message", request)
    if (request.login) {
        await login();
        //setUILogin();
        // await update();
    }
}

async function login() {
    const state = Date.now().toString();
    const redirectUrl = chrome.identity.getRedirectURL();
    console.log("REDIRECT:", redirectUrl);
    const codeVerifier = createRandomString();
    const hash = await sha256(codeVerifier);
    const codeChallenge = bufferToBase64UrlEncoded(hash);

    var launchWebAuthFlow = new Promise((resolve) => {
        chrome.identity.launchWebAuthFlow({
            url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?' +
                `client_id=${adClientId}` +
                '&response_type=code' +
                `&redirect_uri=${redirectUrl}` +
                '&scope=Mail.ReadWrite offline_access' +
                `&state=${state}` +
                '&prompt=select_account' +
                '&code_challenge_method=S256' +
                `&code_challenge=${codeChallenge}`,
            interactive: true
        }, (redirectUrl) => { resolve(redirectUrl) });
    });

    const responseUrl = await launchWebAuthFlow;

    console.log("LOGIN ATTEMPT");

    var params = (new URL(responseUrl)).searchParams;
    if (params.get("state") !== state) {
        console.log("ERROR: states not equal");
        return;
    }
    const code = params.get("code");
    let res = await fetch(tokenEndpoint, {
        method: 'POST',
        body: new URLSearchParams({
            'client_id': adClientId,
            'scope': 'Mail.ReadWrite offline_access',
            'code': code,
            'redirect_uri': redirectUrl,
            'grant_type': 'authorization_code',
            'code_verifier': codeVerifier,
        })
    });

    console.log("finished token attempt");

    if (!res.ok) {
        await console.log(res.json());
        throw new Error('HTTP error: status = ' + res.status);
    }

    res = await res.json();
    console.log("token good?");
    saveToken(res);
    console.log(res);
    setUILogin();
}

function saveToken(res) {
    console.log("token: " + res.access_token);
    chrome.storage.local.get(['token'], function (res) { console.log("check token " + res.token) });
    chrome.storage.local.set({ token: res.access_token });
    chrome.storage.local.set({ refresh_token: res.refresh_token });
    const expiration = new Date(Date.now() + parseInt(res.expires_in) * 1000);
    chrome.storage.local.set({ expiration: expiration.getTime().toString() });
}

async function onAlarm(alarm) {
    await update();
}

async function update() {
    console.log("update!");
    await checkRefreshToken();
    const count = await getUnreadCount();
    setUnreadCount(count);
    //todo: add email displays
}

async function checkRefreshToken() {
    const token = chrome.storage.local.get(['token'], function (res) { return res.token });
    if (token === null) {
        console.log("logged out!");
        setUILogout();
        window.chrome.runtime.sendMessage({ login: false });
    }
    const expiration = new Date(parseInt(chrome.storage.local.get(['expiration'], function (exp) { return exp.expiration })));
    let refreshTime = new Date(expiration.getTime() - refreshWindowMS);
    if (new Date() >= refreshTime) {
        await refreshToken();
    }
}

async function refreshToken() {
    console.log("refresh!");
    let res = await fetch(tokenEndpoint, {
        method: 'POST',
        body: new URLSearchParams({
            'client_id': adClientId,
            'scope': 'Mail.ReadWrite offline_access',
            'refresh_token': chrome.storage.local.get(['refresh_token'], function (res) { return res.refresh_token }),
            'grant_type': 'refresh_token'
        })
    });
    console.log("finished refresh token attempt");

    if (!res.ok) {
        await console.log(res.json());
        throw new Error('HTTP error: status = ' + res.status);
    }

    res = await res.json();
    saveToken(res);
}

function getToken(){
    return new Promise((resolve) => {
        chrome.storage.local.get(['token'], (res) => { resolve(res.token)} );
    })
}

async function getUnreadCount() {
    console.log("checking unread count");
    chrome.storage.local.get(['token'], (res) => {
        console.log(res.token);
        var token = res.token;
        console.log("check token: " + token);
        let res = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox?$select=unreadItemCount', {
            headers: new Headers({
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json'
            })
        });
    });
    

        if (!res.ok) {
            await console.log(res.json());
            throw new Error('HTTP error: status = ' + res.status);
        }

        res = await res.json();
        return res.unreadItemCount;

}

function setUnreadCount(count) {
    chrome.action.setBadgeBackgroundColor({ color: [208, 0, 24, 255] });
    chrome.action.setBadgeText({
        text: count === 0 ? '' : count.toString()
    });
}



//UTILS borrowed from another project
const createRandomString = () => {
    const charset =
        '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_~.';
    let random = '';
    const randomValues = Array.from(
        crypto.getRandomValues(new Uint8Array(43))
    );
    randomValues.forEach(v => (random += charset[v % charset.length]));
    return random;
};

const urlEncodeB64 = (input) => {
    const b64Chars = { '+': '-', '/': '_', '=': '' };
    return input.replace(/[\+\/=]/g, (m) => b64Chars[m]);
};

const bufferToBase64UrlEncoded = input => {
    const ie11SafeInput = new Uint8Array(input);
    return urlEncodeB64(
        btoa(String.fromCharCode(...Array.from(ie11SafeInput)))
    );
};

function sha256(plain) {
    const data = new TextEncoder().encode(plain);
    return crypto.subtle.digest('SHA-256', data);
}
