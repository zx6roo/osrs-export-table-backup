// Tool: Export OSRS Table Backup (multiple pages)
// Description: Exports all saved OSRS Wiki table backup data from localStorage to an Excel (.xlsx) file with proper headers and filename including the current date.
// Allows exporting multiple pages' saved data in one file. Dynamically loads XLSX library if missing.
// Tags: OSRS, wiki, localStorage, bookmarklet, JavaScript, export, Excel, XLSX
// ***DO NOT copy comments into bookmarklet ***

(async () => {
  // Load XLSX library dynamically if not already loaded
  if (typeof XLSX === 'undefined') {
    await new Promise((resolve) => {
      const script = document.createElement('script');
      script.src = 'https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js';
      script.onload = resolve;
      document.head.appendChild(script);
    });
  }

  // Try to load saved backup data from localStorage
  let collected = [];
  try {
    collected = JSON.parse(localStorage.getItem("osrsTableBackup") || "[]");
  } catch (e) {
    alert("❌ Failed to load collected data.");
    return;
  }

  // Alert if no data found
  if (!collected.length) {
    alert("No data found. Did you run the collector on any pages?");
    return;
  }

  // Prepare data array for Excel sheet with headers
  const sheetData = [
    ["Page ID", "Page Name", "URL", "Table No", "Highlight String"],
    ...collected.map(row => [
      row.pageId,
      row.pageName,
      row.url,
      row.tableNo,
      row.highlightString
    ])
  ];

  // Create worksheet and workbook
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "All Pages");

  // Generate filename with current date
  const dateStr = new Date().toISOString().slice(0, 10);
  const filename = `OSRC wiki tables backup ${dateStr}.xlsx`;

  // Write workbook to binary array
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  // Create Blob and download link
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;

  // Append, click to trigger download, then cleanup
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
  }, 100);
})();
