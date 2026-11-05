let wfmChart = null;

document.getElementById("excelFile")
.addEventListener("change", function(event) {

    const file = event.target.files[0];

    if (!file) return;

    document.getElementById("uploadStatus").textContent =
        "Reading Excel file...";

    const reader = new FileReader();

    reader.onload = function(e) {

        try {

            const data = new Uint8Array(e.target.result);

            const workbook = XLSX.read(data, {
                type: "array"
            });

            const sheetName = workbook.SheetNames[0];

            const worksheet = workbook.Sheets[sheetName];

            const rows = XLSX.utils.sheet_to_json(
                worksheet,
                { defval: "" }
            );

            if (rows.length === 0) {
                throw new Error("Excel sheet is empty.");
            }

            const columns = Object.keys(rows[0]);

            document.getElementById("totalRecords")
                .textContent = rows.length;

            document.getElementById("totalColumns")
                .textContent = columns.length;

            document.getElementById("totalSheets")
                .textContent = workbook.SheetNames.length;

            document.getElementById("uploadStatus")
                .textContent =
                "Excel loaded successfully: " + file.name;

            showTable(columns, rows);

            showChart(rows.length);

        } catch(error) {

            document.getElementById("uploadStatus")
                .textContent =
                "Error reading file: " + error.message;

            console.error(error);

        }

    };

    reader.readAsArrayBuffer(file);

});


function showTable(columns, rows) {

    const table = document.getElementById("dataTable");

    const thead = table.querySelector("thead");

    const tbody = table.querySelector("tbody");

    thead.innerHTML = "";

    tbody.innerHTML = "";

    const headerRow = document.createElement("tr");

    columns.forEach(column => {

        const th = document.createElement("th");

        th.textContent = column;

        headerRow.appendChild(th);

    });

    thead.appendChild(headerRow);

    rows.slice(0, 100).forEach(row => {

        const tr = document.createElement("tr");

        columns.forEach(column => {

            const td = document.createElement("td");

            td.textContent = row[column];

            tr.appendChild(td);

        });

        tbody.appendChild(tr);

    });

}


function showChart(totalRecords) {

    const ctx = document
        .getElementById("wfmChart")
        .getContext("2d");

    if (wfmChart) {
        wfmChart.destroy();
    }

    wfmChart = new Chart(ctx, {

        type: "bar",

        data: {

            labels: ["Excel Records"],

            datasets: [{

                label: "Total Records",

                data: [totalRecords]

            }]

        },

        options: {

            responsive: true,

            maintainAspectRatio: false,

            scales: {

                y: {
                    beginAtZero: true
                }

            }

        }

    });

}
