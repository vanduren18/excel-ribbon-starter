/* globals Office, Excel */
Office.onReady(() => {
  // Office is klaar
});

export async function onConnectie(event) {
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
      // Fallback: niets doen
      console.log("Connectie: geen UI beschikbaar.");
    }
  } catch (e) {
    console.error(e);
  } finally {
    event.completed();
  }
}

export async function onDownload(event) {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:C1");
      range.values = [["ID", "Naam", "Waarde"]];
      const data = [
        [1, "Alpha", 10],
        [2, "Beta",  20],
        [3, "Gamma", 30]
      ];
      sheet.getRange("A2:C4").values = data;
      sheet.getUsedRange().format.autofitColumns();
    });
  } catch (e) {
    console.error(e);
  } finally {
    event.completed();
  }
}

export async function onFunctie(event) {
  try {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("values");
      await ctx.sync();
      const vals = range.values.map(row =>
        row.map(cell => (typeof cell === "number" ? cell * 2 : cell))
      );
      range.values = vals;
    });
  } catch (e) {
    console.error(e);
  } finally {
    event.completed();
  }
}

// Zorg dat functies op window staan als modules niet ondersteund zijn
if (typeof window !== "undefined") {
  window.onConnectie = onConnectie;
  window.onDownload  = onDownload;
  window.onFunctie   = onFunctie;
}
