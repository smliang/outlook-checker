function login() {
    console.log("clicked!");
    window.chrome.runtime.sendMessage({login: true});
    window.close();
}

document.addEventListener('DOMContentLoaded', function() {
    document.getElementById("login").addEventListener('click', login);
});

