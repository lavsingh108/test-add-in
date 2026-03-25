Office.onReady((info) => {
    const status = document.getElementById("status");
    if (info.host === Office.HostType.Outlook) {
        status.innerText = "Success! Connected to Outlook.";
    } else {
        status.innerText = "Running outside of Outlook.";
    }
});