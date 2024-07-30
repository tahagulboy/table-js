document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // İlk sayfayı al
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Verileri tabloya aktar
        const htmlString = XLSX.utils.sheet_to_html(worksheet);
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = htmlString;

        const importedTable = tempDiv.querySelector('table');
        importTable(importedTable);
    };
    reader.readAsArrayBuffer(file);
}

function importTable(importedTable) {
    const dataTable = document.getElementById('dataTable');
    
    // Başlıkları ve verileri temizle
    dataTable.innerHTML = '';
    
    // Başlıkları kopyala
    const header = importedTable.querySelector('thead').innerHTML;
    dataTable.innerHTML += `<thead>${header}</thead>`;
    
    // Verileri kopyala
    const body = importedTable.querySelector('tbody').innerHTML;
    dataTable.innerHTML += `<tbody>${body}</tbody>`;
}

document.getElementById('addRow').addEventListener('click', () => {
    const table = document.getElementById('dataTable');
    const rowCount = table.rows.length;
    const colCount = table.rows[0].cells.length;
    const newRow = table.insertRow(rowCount);

    for (let i = 0; i < colCount; i++) {
        const newCell = newRow.insertCell(i);
        newCell.contentEditable = "true";
        newCell.textContent = "Yeni Veri";
    }
});

document.getElementById('addColumn').addEventListener('click', () => {
    const table = document.getElementById('dataTable');
    const rowCount = table.rows.length;

    for (let i = 0; i < rowCount; i++) {
        const row = table.rows[i];
        const newCell = row.insertCell(-1);
        newCell.contentEditable = "true";
        if (i === 0) {
            newCell.textContent = "Yeni Başlık";
        } else {
            newCell.textContent = "Yeni Veri";
        }
    }
});

document.getElementById('deleteRow').addEventListener('click', () => {
    const table = document.getElementById('dataTable');
    const rowCount = table.rows.length;

    if (rowCount > 1) { // Başlık satırını silme
        table.deleteRow(rowCount - 1);
    } else {
        alert("Tabloda silinecek satır yok.");
    }
});

document.getElementById('deleteColumn').addEventListener('click', () => {
    const table = document.getElementById('dataTable');
    const colCount = table.rows[0].cells.length;

    if (colCount > 1) { // İlk sütunu silme
        for (let i = 0; i < table.rows.length; i++) {
            table.rows[i].deleteCell(-1);
        }
    } else {
        alert("Tabloda silinecek sütun yok.");
    }
});

document.getElementById('saveButton').addEventListener('click', () => {
    const table = document.getElementById('dataTable');
    const wb = XLSX.utils.table_to_book(table, {sheet: "Sheet1"});
    const wbout = XLSX.write(wb, {bookType: 'xlsx', type: 'binary'});

    function s2ab(s) {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    const blob = new Blob([s2ab(wbout)], {type: "application/octet-stream"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "tablo.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
});
