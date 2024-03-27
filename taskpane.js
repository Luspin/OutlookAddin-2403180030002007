Office.onReady(async (info) => {
    console.log("Office.onReady called");

    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook detected");
    }
});


