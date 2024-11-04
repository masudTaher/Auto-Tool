document.getElementById('loadButton').addEventListener('click', handleFile);
document.getElementById('addRowButton').addEventListener('click', addRow);
document.getElementById('saveButton').addEventListener('click', saveFile);

let workbook;
let fileName = ''; // Standardname für die Datei

function handleFile(event) {
    const file = document.getElementById('fileInput').files[0];
    const reader = new FileReader();

    if (!file) {
        alert("Bitte wähle eine Datei aus.");
        return;
    }

    fileName = file.name; // Speichere den Namen der hochgeladenen Datei

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        displayData(json);
    };

    reader.readAsArrayBuffer(file);
}

function displayData(data) {
    const table = document.getElementById('dataTable');
    table.innerHTML = '';

    data.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');

        row.forEach((cell) => {
            const td = document.createElement('td');
            td.textContent = cell;
            td.contentEditable = true; // Zellen bearbeitbar machen
            tr.appendChild(td);
        });

        table.appendChild(tr);
    });
}

function addRow() {
    const table = document.getElementById('dataTable');
    const tr = document.createElement('tr');

    for (let i = 0; i < 3; i++) {
        const td = document.createElement('td');
        td.contentEditable = true; // Neue Zellen ebenfalls bearbeitbar
        tr.appendChild(td);
    }

    table.appendChild(tr);
}

function saveFile() {
    const table = document.getElementById('dataTable');
    const data = [];

    // Extrahiere Daten aus der Tabelle
    for (let row of table.rows) {
        const rowData = [];
        for (let cell of row.cells) {
            rowData.push(cell.textContent);
        }
        data.push(rowData);
    }

    // Erstelle ein neues Arbeitsblatt und Workbook
    const newWorksheet = XLSX.utils.aoa_to_sheet(data);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');

    // Speichere die Datei mit dem ursprünglichen Namen
    if (fileName) {
        XLSX.writeFile(newWorkbook, fileName);
        // Automatisches Neuladen der Tabelle mit den neuen Daten
        displayData(data); // Aktualisiere die Tabelle mit den gespeicherten Daten
    } else {
        alert("Bitte lade zuerst eine Datei hoch.");
    }
}
