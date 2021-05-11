//borrowed from another project

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
    const b64Chars = {'+': '-', '/': '_', '=': ''};
    return input.replace(/[\+\/=]/g, (m) => b64Chars[m]);
};

const bufferToBase64UrlEncoded = input => {
    const ie11SafeInput = new Uint8Array(input);
    return urlEncodeB64(
        window.btoa(String.fromCharCode(...Array.from(ie11SafeInput)))
    );
};

function sha256(plain) {
    const data = new TextEncoder().encode(plain);
    return crypto.subtle.digest('SHA-256', data);
}
