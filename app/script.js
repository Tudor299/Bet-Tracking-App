function openTeamsStats() {
    window.open('view_table.html?file=football_statistics.xlsx', '_blank');
}

async function refreshPlacedBets() {
    const button = document.getElementById("refresh-bets-button");
    button.disabled = true;
    button.textContent = "Processing...";
    const response = await fetch("http://127.0.0.1:8000/refresh_bets");
    const data = await response.json();
    button.disabled = false;
    button.textContent = "Refresh";
}

function openPlacedBets() {
    window.open('view_table.html?file=placed_bets.xlsx', '_blank')
}

async function refreshTeamsStats() {
    const button = document.getElementById("refresh-teams-stats-button");
    button.disabled = true;
    button.textContent = "Processing...";
    const response = await fetch("http://127.0.0.1:8000/refresh_teams");
    const data = await response.json();
    button.disabled = false;
    button.textContent = "Refresh";
}

function getQueryParam(name) {
    const params = new URLSearchParams(window.location.search);
    return params.get(name);
}

fetch('placed_bets.xlsx')
    .then(res => res.arrayBuffer())
    .then(buffer => {
        const workbook = XLSX.read(buffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const won_perc = sheet['N2'];
        const value_won_perc = won_perc.v ? won_perc.v : 'Cell not found';
        const won = sheet['L2'];
        const value_won = won ? won.v : 'Cell not found';
        const balance = sheet['T2'];
        const value_balance = balance ? balance.v : 'Cell not found';

        document.getElementById('win-rate-value-label').textContent = value_won_perc + "%";
        document.getElementById('won-value-value').textContent = value_won;
        document.getElementById('balance-value-label').textContent = value_balance + "RON";
    })
    .catch(err => {
        document.getElementById('cellValue').textContent = 'Error reading Excel file';
        console.error(err);
    });

const fileName = getQueryParam('file');

if (fileName) {
    fetch(fileName)
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const outputDiv = document.getElementById('output');

            const sheet1 = workbook.SheetNames[0];
            const html1 = XLSX.utils.sheet_to_html(workbook.Sheets[sheet1]);
            outputDiv.innerHTML = html1;

            if (workbook.SheetNames.length >= 2) {
                const sheet2 = workbook.SheetNames[1];
                const html2 = XLSX.utils.sheet_to_html(workbook.Sheets[sheet2]);
                outputDiv.innerHTML += html2;
            }

        })
        .catch(err => {
            document.getElementById('output').textContent = 'Error loading Excel file.';
            console.error(err);
        });
} else {
    document.getElementById('output').textContent = 'No file specified.';
}