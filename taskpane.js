/* global document, Office, Word, alert */

// Immediate debug check to confirm the script file is loading
console.log("Script file successfully loaded by the browser engine!");

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const scoreButton = document.getElementById("run");
    if (scoreButton) {
      scoreButton.addEventListener("click", () => {
        alert("Button clicked! Attempting to execute calculations...");
        tryCatch(run);
      });
    } else {
      console.error("Fatal Error: Could not find HTML element with id='run'");
    }
  }
});

async function run() {
  await Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();

    const getVal = (cellText) => {
      if (!cellText) return 0;
      const num = parseFloat(cellText.replace(/[^\d.-]/g, ""));
      return isNaN(num) ? 0 : num;
    };

    const idx1 = 1;
    const idx2 = 2;
    const idx3 = 3;
    const idx4 = 4;
    const idx5 = 5;

    for (let i = 0; i < tables.items.length; i++) {
      const tbl = tables.items[i];
      const rows = tbl.rows;
      rows.load("items");
      await context.sync();

      if (rows.items.length < 5) continue;

      const row1 = rows.items[idx1];
      const row2 = rows.items[idx2];
      const row3 = rows.items[idx3];
      const row4 = rows.items[idx4];

      row1.cells.load("items");
      row2.cells.load("items");
      row3.cells.load("items");
      row4.cells.load("items");
      await context.sync();

      if (row1.cells.items.length < 6 || 
          row2.cells.items.length < 6 || 
          row3.cells.items.length < 6 || 
          row4.cells.items.length < 6) {
        continue;
      }

      for (let c = 1; c <= 5; c++) {
        row1.cells.items[c].body.load("text");
        row2.cells.items[c].body.load("text");
        row3.cells.items[c].body.load("text");
        row4.cells.items[c].body.load("text");
      }
      await context.sync();

      try {
        const valB2 = getVal(row1.cells.items[idx1].body.text);
        const valC2 = getVal(row1.cells.items[idx2].body.text);
        const valD2 = getVal(row1.cells.items[idx3].body.text);
        const valE2 = getVal(row1.cells.items[idx4].body.text);

        const valB3 = valB2 * 3;
        const valC3 = valC2 * 3;
        const valD3 = valD2 * 2;
        const valE3 = valE2 * 2;

        row2.cells.items[idx1].body.insertText(valB3 === 0 ? "" : valB3.toString(), "Replace");
        row2.cells.items[idx2].body.insertText(valC3 === 0 ? "" : valC3.toString(), "Replace");
        row2.cells.items[idx3].body.insertText(valD3 === 0 ? "" : valD3.toString(), "Replace");
        row2.cells.items[idx4].body.insertText(valE3 === 0 ? "" : valE3.toString(), "Replace");

        const valF3 = valB3 + valC3 + valD3 + valE3;
        row2.cells.items[idx5].body.insertText(valF3 === 0 ? "" : valF3.toString(), "Replace");

        const valB4 = getVal(row3.cells.items[idx1].body.text);
        const valC4 = getVal(row3.cells.items[idx2].body.text);
        const valD4 = getVal(row3.cells.items[idx3].body.text);
        const valE4 = getVal(row3.cells.items[idx4].body.text);

        const valF4 = valB4 * 3 + valC4 * 3 + valD4 * 2 + valE4 * 2;
        row3.cells.items[idx5].body.insertText(valF4 === 0 ? "" : valF4.toString(), "Replace");

        const valB5 = getVal(row4.cells.items[idx1].body.text);
        const valC5 = getVal(row4.cells.items[idx2].body.text);
        const valD5 = getVal(row4.cells.items[idx3].body.text);
        const valE5 = getVal(row4.cells.items[idx4].body.text);

        const valF5 = valB5 * 3 + valC5 * 3 + valD5 * 2 + valE5 * 2;
        row4.cells.items[idx5].body.insertText(valF5 === 0 ? "" : valF5.toString(), "Replace");

      } catch (error) {
        alert("Table execution failed: " + error.message);
      }
    }

    const docFields = context.document.body.fields;
    docFields.load("items");
    await context.sync();

    for (let f = 0; f < docFields.items.length; f++) {
      docFields.items[f].updateResult();
    }
    await context.sync();

    alert("Scores successfully updated!");
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    alert("Runtime Error: " + error.message);
  }
}
