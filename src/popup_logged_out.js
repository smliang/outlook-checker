function login() {
    console.log("clicked!");
    chrome.runtime.sendMessage({login: true});
   // setTimeout(window.close(), 2000);
}

document.addEventListener('DOMContentLoaded', function() {
    document.getElementById("login").addEventListener('click', login);
});

