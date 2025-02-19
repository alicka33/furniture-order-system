document.addEventListener("DOMContentLoaded", function() {
    const table = document.getElementById("dynamicTable");
    const tbody = table.getElementsByTagName("tbody")[0];
    const totalSumCell = document.getElementById("totalSum");
    const generateCSVBtn = document.getElementById("generateCSVBtn");
    const importExcelBtn = document.getElementById("importExcelBtn");

    // Include the script that handles Excel import
    const script = document.createElement('script');
    script.src = 'importExcel.js';  // Make sure this path is correct
    document.head.appendChild(script);

    // Load materials from JSON file
    function loadMaterials() {
        fetch('materials.json')
            .then(response => response.json())
            .then(data => {
                // Create and add datalist with options
                createDatalist(data.materials);
                // Initialize the first row with material options
                setupMaterialInputs();
            })
            .catch(error => console.error('Error loading materials:', error));
    }

    // Create a datalist element with options
    function createDatalist(materials) {
        const datalist = document.getElementById('materialsList');
        materials.forEach(material => {
            const option = document.createElement("option");
            option.value = material.value;
            option.textContent = material.label;
            datalist.appendChild(option);
        });
    }

    // Add a new row to the table
    function addRow() {
        const newRow = tbody.insertRow();
        for (let i = 0; i < table.rows[0].cells.length - 1; i++) { // -1 for the "Akcje" column
            const newCell = newRow.insertCell(i);
            if (i === 1) { // For the "Materiał" column
                const input = document.createElement('input');
                input.type = 'text';
                input.className = 'materialInput';
                input.setAttribute('list', 'materialsList'); // Set list attribute
                input.placeholder = 'Wpisz materiał...';
                newCell.appendChild(input);
            } else if (i < 10) { // Editable cells for D2 and S2
                newCell.contentEditable = "true";
                newCell.addEventListener('input', calculateRow);
                newCell.addEventListener('keydown', handleKeyDown);
            }
        }
        // Add an empty cell for the "Razem" column
        newRow.insertCell(table.rows[0].cells.length - 2);
        
        // Create and append the delete button
        const actionCell = newRow.insertCell(table.rows[0].cells.length - 1); // Last
        const deleteBtn = document.createElement("button");
        deleteBtn.textContent = "Usuń";
        deleteBtn.className = "deleteBtn";
        deleteBtn.addEventListener("click", function() {
            if (tbody.rows.length > 1) { // Ensure there's at least one row to keep
                newRow.remove();
                calculateTotalSum();
            }
        });
        actionCell.appendChild(deleteBtn);

        // Setup the new material input field
        setupMaterialInputs();
    }

    // Calculate values in a row
    function calculateRow(e) {
        const row = e.target.closest('tr');
        const ilosc = parseFloat(row.cells[2].innerText.replace(',', '.')) || 0;
        const grubosc = parseFloat(row.cells[3].innerText.replace(',', '.')) || 0;
        const dlugosc = parseFloat(row.cells[4].innerText.replace(',', '.')) || 0;
        const szerokosc = parseFloat(row.cells[5].innerText.replace(',', '.')) || 0;
        // const d1 = parseFloat(row.cells[6].innerText.replace(',', '.')) || 0;
        // const d2 = parseFloat(row.cells[7].innerText.replace(',', '.')) || 0;
        // const s1 = parseFloat(row.cells[8].innerText.replace(',', '.')) || 0;
        // const s2 = parseFloat(row.cells[9].innerText.replace(',', '.')) || 0;
        const d1 = row.cells[6].innerText.trim();
        const d2 = row.cells[7].innerText.trim();
        const s1 = row.cells[8].innerText.trim();
        const s2 = row.cells[9].innerText.trim();

        const plytaM2 = (dlugosc * szerokosc) / 1000000;
        row.cells[10].innerText = plytaM2.toFixed(3);

        const faktor = 1.2;
        row.cells[11].innerText = faktor;

        const cenaM2 = grubosc * 0.06;  // Example calculation
        row.cells[12].innerText = cenaM2.toFixed(2);

        const d1Value = d1 ? 1 : 0;
        const d2Value = d2 ? 1 : 0;
        const s1Value = s1 ? 1 : 0;
        const s2Value = s2 ? 1 : 0;

        // Obliczenie obzezeMb z uwzględnieniem warunków
        const obzezeMb = (dlugosc * d1Value + dlugosc * d2Value + szerokosc * s1Value + szerokosc * s2Value) / 1000;
        //const obzezeMb = (dlugosc * d1 + dlugosc * d2 + szerokosc * s1 + szerokosc * s2) / 1000;
        row.cells[13].innerText = obzezeMb.toFixed(2);


        const cenaMb = grubosc / 160;  // Example calculation
        row.cells[14].innerText = cenaMb.toFixed(2);

        const materialInput = row.cells[1].querySelector('.materialInput');
        const selectedMaterial = materialInput ? materialInput.value : '';
        const materialPrice = parseFloat(selectedMaterial) || 0;

        const material = plytaM2 * faktor * cenaM2 + obzezeMb * cenaMb;
        row.cells[15].innerText = material.toFixed(2); // Materialx

        const format = 2 * (dlugosc + szerokosc) / 1000 * 1.2;
        row.cells[16].innerText = format.toFixed(2); // Format

        const obrzeze = grubosc * obzezeMb * cenaMb;
        row.cells[17].innerText = obrzeze.toFixed(2);

        const razem = ilosc * (material + format + obrzeze);
        row.cells[18].innerText = razem.toFixed(2);

        calculateTotalSum();
    }

    // Calculate total sum of "Razem" column
    function calculateTotalSum() {
        let totalSum = 0;
        for (let i = 0; i < tbody.rows.length; i++) {
            const razem = parseFloat(tbody.rows[i].cells[18].innerText) || 0;
            totalSum += razem;
        }
        totalSumCell.innerText = totalSum.toFixed(2);
    }

       // Handle keydown events for navigating cells and adding rows
       function handleKeyDown(e) {
        const key = e.key;
        const target = e.target;
        const row = target.closest('tr');
        const cellIndex = target.cellIndex;
        const rowIndex = row.rowIndex - 1;

        if (key === "ArrowRight" && cellIndex < row.cells.length - 1) {
            row.cells[cellIndex + 1].focus();
            e.preventDefault();
        } else if (key === "ArrowLeft" && cellIndex > 0) {
            row.cells[cellIndex - 1].focus();
            e.preventDefault();
        } else if (key === "ArrowDown" && rowIndex < tbody.rows.length - 1) {
            tbody.rows[rowIndex + 1].cells[cellIndex].focus();
            e.preventDefault();
        } else if (key === "ArrowUp" && rowIndex > 0) {
            tbody.rows[rowIndex - 1].cells[cellIndex].focus();
            e.preventDefault();
        } else if (key === "Enter") {
            if (target.contentEditable === "true") {
                addRow();
                tbody.rows[rowIndex + 1].cells[cellIndex].focus();
            }
            e.preventDefault();
        }
    }

    // Add event listeners to editable cells
    function setupEditableCells() {
        table.querySelectorAll('td[contenteditable="true"]').forEach(cell => {
            cell.addEventListener('keydown', handleKeyDown);
        });
    }

    // Add event listeners to material input fields
    function setupMaterialInputs() {
        table.querySelectorAll('input.materialInput').forEach(input => {
            input.addEventListener('input', calculateRow);
            input.addEventListener('change', calculateRow);
            // Ensure list attribute is set for each input
            if (!input.hasAttribute('list')) {
                input.setAttribute('list', 'materialsList');
            }
        });
    }

    setupEditableCells();
    loadMaterials(); // Load materials and initialize datalist
    setupMaterialInputs();

    // Recalculate row values on blur event
    table.addEventListener("blur", function(e) {
        if (e.target.matches('td[contenteditable="true"]')) {
            calculateRow(e);
        }
    }, true);

    // Add event listener for dynamically added delete buttons
    tbody.addEventListener('click', function(e) {
        if (e.target && e.target.classList.contains('deleteBtn')) {
            const row = e.target.closest('tr');
            if (tbody.rows.length > 1) { // Ensure there's at least one row to keep
                row.remove();
                calculateTotalSum();
            }
        }
    });

    // Initial call to setup the table with one row
    addRow();

    function removePolishCharacters(str) {
        const polishChars = {
            'ą': 'a', 'ć': 'c', 'ę': 'e', 'ł': 'l', 'ń': 'n', 'ó': 'o', 'ś': 's', 'ź': 'z', 'ż': 'z',
            'Ą': 'A', 'Ć': 'C', 'Ę': 'E', 'Ł': 'L', 'Ń': 'N', 'Ó': 'O', 'Ś': 'S', 'Ź': 'Z', 'Ż': 'Z'
        };
        return str.split('').map(char => polishChars[char] || char).join('');
    }

    function generateCSV() {
        let csv = [];
        
        // Add three empty rows at the beginning
        csv.push("");
        csv.push("");
        csv.push("");
    
        // Define headers
        const headers = [
            "", "Lp.", "Nr rysunku", "Nazwa komponentu", // New initial columns
            "Nazwa elementu", "Material", "Ilosc", "Grubosc", "Dlugosc", "Szerokosc", // Existing columns up to "Szerokość"
            "", "", // Two empty columns
            "Obrzeze (dol)", "Obrzeze (gora)", "Obrzeze (przod)", "Obrzeze (tyl)", // Existing columns from "Obrzeże (dół)" to "Obrzeże (tył)"
            "<D", "<S", "Nr. proj.", "Projekt" // New final columns
        ];
        csv.push(headers.join(";")); // Add headers to CSV
    
        // Iterate through table rows
        tbody.querySelectorAll('tr').forEach((row, rowIndex) => {
            const cells = [];
    
            // Add four empty cells at the beginning for the new columns
            cells.push(""); // Empty column
            cells.push((rowIndex + 1) + "."); // Lp. column with row number

            // Add empty cells for "Nr rysunku" and "Nazwa komponentu"
            cells.push("");
            cells.push("");
    
            // Extract values for existing columns up to "Szerokość"
            Array.from(row.querySelectorAll('td')).forEach((td, index) => {
                if (index <= 5) { // Up to "Szerokość"
                    let value = td.innerText.trim();
    
                    if (index === 1) { // Handle input fields in the "Materiał" column
                        const input = td.querySelector('input.materialInput');
                        if (input) {
                            value = input.value.trim(); // Use the input value for "Materiał"
                        }
                    }
    
                    if (index === 0 || index === 1) { // Remove Polish characters for the first two columns
                        value = removePolishCharacters(value);
                    }
    
                    if (index > 1 && index < 6) { // Replace dots with commas in specific numeric fields
                        value = value.replace('.', ',');
                    }
    
                    cells.push(value);
                }
            });
    
            // Add two empty cells after "Szerokość"
            cells.push("", "");
    
            // Extract values for "Obrzeże (dół)" to "Obrzeże (tył)"
            Array.from(row.querySelectorAll('td')).forEach((td, index) => {
                if (index >= 6 && index <= 9) {
                    let value = td.innerText.trim();
                    cells.push(value);
                }
            });
    
            // Add four empty cells for "<D", "<S", "Nr. proj.", "Projekt"
            cells.push("", "", "", "");
    
            csv.push(cells.join(";")); // Add row to CSV
        });
    
        // Add total sum row if needed
        // Add "Opakowanie" row
        //csv.push(["", "", "", "", "Opakowanie", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""].join(";"));
    
        // Add "Podsumowanie" row with total sum
        // csv.push(["", "", "", "", "Podsumowanie", "", "", "", "", "", "", "", "", "", "", "", "", "", "", totalSumCell.innerText.replace('.', ',')].join(";"));
        
        // Download CSV file
        downloadCSV(csv.join("\n"));
    }    
    

    // Function to trigger download of CSV file
    function downloadCSV(csv) {
        // Pobierz tytuł tabeli
        const tableTitle = document.getElementById("tableTitle").innerText.trim();
        
        // Zamień znaki niedozwolone w nazwie pliku
        const safeTitle = tableTitle.replace(/[^a-zA-Z0-9_-]/g, "_");
    
        // Utwórz plik CSV
        const csvFile = new Blob([csv], { type: "text/csv" });
        const downloadLink = document.createElement("a");
    
        // Ustaw nazwę pliku na podstawie tytułu tabeli
        downloadLink.download = `${safeTitle}.csv`;
        downloadLink.href = window.URL.createObjectURL(csvFile);
        downloadLink.style.display = "none";
    
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
    }

    // Add event listener to the CSV button
    generateCSVBtn.addEventListener("click", generateCSV);  
    
    importExcelBtn.addEventListener('change', function(event) {
        const file = event.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
                // Clear existing rows  // to się jeszcze zobaczy 
                //tbody.innerHTML = "";
    
                // Populate table with imported data
                jsonData.forEach((row, rowIndex) => {
                    if (rowIndex > 0) { // Skip header row
                        const newRow = tbody.insertRow();
    
                        // Assuming columns are in a fixed order
                        const nazwa = row[0] || '';
                        const mat = row[1] || '';
                        const ilosc = parseFloat(row[2]) || 0;
                        const grubosc = parseFloat(row[3]) || 0;
                        const dlugosc = parseFloat(row[4]) || 0;
                        const szerokosc = parseFloat(row[5]) || 0;
                        
                        const d1 = row[6] !== undefined && row[6] !== null ? String(row[6]).trim() : "";
                        const d2 = row[7] !== undefined && row[7] !== null ? String(row[7]).trim() : "";
                        const s1 = row[8] !== undefined && row[8] !== null ? String(row[8]).trim() : "";
                        const s2 = row[9] !== undefined && row[9] !== null ? String(row[9]).trim() : "";

                        // const d1 = parseFloat(row[6]) || 0;
                        // const d2 = parseFloat(row[7]) || 0;
                        // const s1 = parseFloat(row[8]) || 0;
                        // const s2 = parseFloat(row[9]) || 0;

                        // Create cells and set values
                        [nazwa, mat, ilosc, grubosc, dlugosc, szerokosc, d1, d2, s1, s2].forEach((value, index) => {
                            const newCell = newRow.insertCell(index);
                            newCell.contentEditable = "true";
                            newCell.innerText = value;
                            newCell.addEventListener('input', calculateRow);
                            newCell.addEventListener('keydown', handleKeyDown);
                        });

                        const plytaM2 = (dlugosc * szerokosc) / 1000000;
                
                        const faktor = 1.2;                        
                        const cenaM2 = grubosc * 0.06;  // Example calculation
                        //row.cells[11].innerText = cenaM2.toFixed(2);
                        
                        
                        const d1Value = d1 ? 1 : 0;
                        const d2Value = d2 ? 1 : 0;
                        const s1Value = s1 ? 1 : 0;
                        const s2Value = s2 ? 1 : 0;

                        // Obliczenie obzezeMb z uwzględnieniem warunków
                        const obzezeMb = (dlugosc * d1Value + dlugosc * d2Value + szerokosc * s1Value + szerokosc * s2Value) / 1000;
                        //const obzezeMb = (dlugosc * 1 + dlugosc * 1 + szerokosc * 1 + szerokosc * 1) / 1000;
                
                        const cenaMb = grubosc / 160;  // Example calculation
                        //row.cells[13].innerText = cenaMb.toFixed(2)
                
                        const material = plytaM2 * faktor * cenaM2 + obzezeMb * cenaMb;
                
                        const format = 2 * (dlugosc + szerokosc) / 1000 * 1.2;
                
                        const obrzeze = grubosc * obzezeMb * cenaMb;
                
                        const razem = ilosc * (material + format + obrzeze);
                        
                        // Dodawanie nowych komórek z wynikami
                        const resultValues = [plytaM2, faktor, cenaM2, obzezeMb, cenaMb, material, format, obrzeze, razem];
                        resultValues.forEach((value, index) => {
                            const newCell = newRow.insertCell(10 + index); // Zaczynamy od 10. kolumny
                            newCell.contentEditable = "false";
                            newCell.innerText = value.toFixed(2);
                        });

                        // Add action cell with delete button
                        addDeleteButton(newRow);
                        calculateRow({ target: newRow });

                    }
                });
    
                // Recalculate totals after importing
                calculateTotalSum();
            };
            reader.readAsArrayBuffer(file);
        }
    });
    
    function addDeleteButton(row) {
        const actionCell = row.insertCell(-1); // Insert at the end
        const deleteBtn = document.createElement("button");
        deleteBtn.textContent = "Usuń";
        deleteBtn.className = "deleteBtn";
        deleteBtn.addEventListener("click", function() {
            if (tbody.rows.length > 1) { // Ensure there's at least one row to keep
                row.remove();
                calculateTotalSum();
            }
        });
        actionCell.appendChild(deleteBtn);
    }
        
    
});

