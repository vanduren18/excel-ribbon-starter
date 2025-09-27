/* globals Office, Excel */
Office.onReady(() => {
  console.log("commands.js geladen en Office is ready");
});

function onConnectie(event) {
  try {
    if (Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync("about:blank", { height: 20, width: 30, displayInIframe: true },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const dlg = asyncResult.value;
            try { dlg.messageParent("Connectietest geslaagd."); } catch(e){}
            setTimeout(() => dlg.close(), 1200);
          }
        });
    } else {
      console.log("Connectie: geen UI beschikbaar.");
    }
  } catch (e) {
    console.error(e);
  } finally {
    if (event && event.completed) event.completed();
  }
}

function onDownload(event) {
  Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("A1:C1").values = [["ID","Naam","Waarde"]];
    sheet.getRange("A2:C4").values = [[1,"Alpha",10],[2,"Beta",20],[3,"Gamma",30]];
    sheet.getUsedRange().format.autofitColumns();
  }).catch(console.error).finally(() => {
    if (event && event.completed) event.completed();
  });
}

function onFunctie(event) {
  Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load("values");
    await ctx.sync();
    const changed = range.values.map(r => r.map(c => (typeof c === "number" ? c * 2 : c)));
    range.values = changed;
  }).catch(console.error).finally(() => {
    if (event && event.completed) event.completed();
  });
}

// Maak functies globaal (vereist voor ExecuteFunction)
if (typeof window !== "undefined") {
  window.onConnectie = onConnectie;
  window.onDownload  = onDownload;
  window.onFunctie   = onFunctie;
}
