/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Correctly locate the DOM button and bind the click event listener 
    // within the active Office lifecycle context.
    const scoreButton = document.getElementById("run");
    if (scoreButton) {
      scoreButton.addEventListener("click", () => tryCatch(run));
    }
  }
});

async function run() {
  await Word.run(async (context) => {
    // Get all tables in the document
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();

    // Helper function to safely extract numeric values from raw cell text
    const getVal = (cellText) => {
      if (!cellText) return 0;
      const num = parseFloat(cellText.replace(/[^\d.-]/g, ""));
      return isNaN(num) ? 0 : num;
    };

    // Explicit indices to prevent text compilation bugs
    const idx1 = 1;
    const idx2 = 2;
    const idx3 = 3;
    const idx4 = 4;
    const idx5 = 5;

    // Loop through every table found in the document
    for (let i = 0; i < tables.items.length; i++) {
      const tbl = tables.items[i];
      const rows = tbl.rows;
      rows.load("items");
      await context.sync();

      // Skip table if it doesn't have at least 5 rows
      if (rows.items.length < 5) continue;

      // Load cells for the rows we need to inspect
      const row1 = rows.items[idx1];
      const row2 = rows.items[idx2];
      const row3 = rows.items[idx3];
      const row4 = rows.items[idx4];

      row1.cells.load("items");
      row2.cells.load("items");
      row3.cells.load("items");
      row4.cells.load("items");
      await context.sync();

      // Skip if any target row lacks required columns
      if (row1.cells.items.length < 6 || 
          row2.cells.items.length < 6 || 
          row3.cells.items.length < 6 || 
          row4.cells.items.length < 6) {
        continue;
      }

      // Deep load the cell body text for our calculations (Columns 1 to 5)
      for (let c = 1; c <= 5; c++) {
        row1.cells.items[c].body.load("text");
        row2.cells.items[c].body.load("text");
        row3.cells.items[c].body.load("text");
        row4.cells.items[c].body.load("text");
      }
      await context.sync();

      try {
        // --- 1. Calculate and update Row 3 (B3, C3, D3, E3) ---
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

        // --- 2. Calculate F3 (Sum of Row 3) ---
        const valF3 = valB3 + valC3 + valD3 + valE3;
        row2.cells.items[idx5].body.insertText(valF3 === 0 ? "" : valF3.toString(), "Replace");

        // --- 3. Calculate and update F4 (Row 4) ---
        const valB4 = getVal(row3.cells.items[idx1].body.text);
        const valC4 = getVal(row3.cells.items[idx2].body.text);
        const valD4 = getVal(row3.cells.items[idx3].body.text);
        const valE4 = getVal(row3.cells.items[idx4].body.text);

        const valF4 = valB4 * 3 + valC4 * 3 + valD4 * 2 + valE4 * 2;
        row3.cells.items[idx5].body.insertText(valF4 === 0 ? "" : valF4.toString(), "Replace");

        // --- 4. Calculate and update F5 (Row 5) ---
        const valB5 = getVal(row4.cells.items[idx1].body.text);
        const valC5 = getVal(row4.cells.items[idx2].body.text);
        const valD5 = getVal(row4.cells.items[idx3].body.text);
        const valE5 = getVal(row4.cells.items[idx4].body.text);

        const valF5 = valB5 * 3 + valC5 * 3 + valD5 * 2 + valE5 * 2;
        row4.cells.items[idx5].body.insertText(valF5 === 0 ? "" : valF5.toString(), "Replace");

      } catch (error) {
        console.error("Error processing table index " + i + ": ", error);
      }
    }

    // Update native document fields (e.g. Page Numbers, TOC)
    const docFields = context.document.body.fields;
    docFields.load("items");
    await context.sync();

    for (let f = 0; f < docFields.items.length; f++) {
      docFields.items[f].updateResult();
    }
    await context.sync();

    console.log("Scores successfully updated!");
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
