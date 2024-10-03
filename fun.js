document.getElementById('fetch-headers').addEventListener('click', function() {
    const url = document.getElementById('excel-url').value;
    if (!url) {
        alert("Please enter a valid Excel file URL.");
        return;
    }

    fetchExcelFile(url, function(workbook) {
        const sheetNames = workbook.SheetNames;
        if (sheetNames.length > 0) {
            displayHeaders(workbook.Sheets[sheetNames[0]]);
        }
    });
});

document.getElementById('fetch-sheet').addEventListener('click', function() {
    const headerChoice = document.getElementById('header-choice').value;
    if (!headerChoice) {
        alert("Please select a header.");
        return;
    }

    const url = document.getElementById('excel-url').value;
    fetchExcelFile(url, function(workbook) {
        const selectedSheet = workbook.Sheets[headerChoice];
        if (selectedSheet) {
            displaySheetContents(selectedSheet);
        } else {
            alert("Could not find the selected sheet.");
        }
    });
});

function fetchExcelFile(url, callback) {
    fetch(url)
        .then(res => res.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            callback(workbook);
        })
        .catch(err => {
            console.error(err);
            alert("Failed to load Excel file. Please check the URL.");
        });
}

function displayHeaders(sheet) {
    const headers = getSheetHeaders(sheet);
    const headerSelect = document.getElementById('header-choice');
    headerSelect.innerHTML = '';

    headers.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.text = header;
        headerSelect.appendChild(option);
    });

    document.getElementById('header-selection').style.display = 'block';
}

function getSheetHeaders(sheet) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const headers = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: C })];
        headers.push(cell ? cell.v : `Column ${C + 1}`);
    }
    return headers;
}

function displaySheetContents(sheet) {
    const jsonSheet = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const contentDiv = document.getElementById('sheet-content');
    contentDiv.innerHTML = JSON.stringify(jsonSheet, null, 4);
}
