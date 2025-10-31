class ExcelProcessor {
    constructor() {
        this.workbook = null;
        this.processedWorkbook = null;
        this.init();
    }

    init() {
        // DOM elements
        this.fileInput = document.getElementById('fileInput');
        this.dropZone = document.getElementById('dropZone');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.status = document.getElementById('status');

        // Event listeners
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        this.dropZone.addEventListener('click', () => this.fileInput.click());
        this.dropZone.addEventListener('dragover', (e) => this.handleDragOver(e));
        this.dropZone.addEventListener('drop', (e) => this.handleFileDrop(e));
        this.downloadBtn.addEventListener('click', () => this.downloadFile());

        this.updateStatus('Please upload an Excel file to begin processing', 'info');
    }

    handleDragOver(e) {
        e.preventDefault();
        this.dropZone.style.borderColor = '#0056b3';
        this.dropZone.style.backgroundColor = '#e3f2fd';
    }

    handleFileDrop(e) {
        e.preventDefault();
        this.dropZone.style.borderColor = '#007cba';
        this.dropZone.style.backgroundColor = '#f8f9fa';
        
        const files = e.dataTransfer.files;
        if (files.length) {
            this.loadFile(files[0]);
        }
    }

    handleFileSelect(e) {
        const files = e.target.files;
        if (files.length) {
            this.loadFile(files[0]);
        }
    }

    loadFile(file) {
        if (!file.name.match(/\.(xlsx|xls|csv)$/)) {
            this.updateStatus('❌ Please upload a valid Excel file (.xlsx, .xls, .csv)', 'error');
            return;
        }

        this.updateStatus('⏳ Loading file...', 'info');
        this.dropZone.innerHTML = '<h3>⏳ Processing...</h3><p>Please wait</p>';

        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                this.workbook = XLSX.read(data, { type: 'array' });
                this.updateStatus('✅ File loaded! Processing...', 'info');
                
                // Automatically start processing
                setTimeout(() => this.processFile(), 500);
                
            } catch (error) {
                this.updateStatus('❌ Error reading file: ' + error.message, 'error');
                this.resetUploadArea();
            }
        };

        reader.onerror = () => {
            this.updateStatus('❌ Error reading file', 'error');
            this.resetUploadArea();
        };

        reader.readAsArrayBuffer(file);
    }
    
    addMultiplicationFormulas(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Kolom O adalah index 14 (0-based)
        const targetColumn = 14;
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cellD = XLSX.utils.encode_cell({ r: R, c: 3 }); // Kolom D
            const cellO = XLSX.utils.encode_cell({ r: R, c: targetColumn }); // Kolom O
            
            // Periksa apakah kolom D tidak kosong
            if (worksheet[cellD] && worksheet[cellD].v && String(worksheet[cellD].v).trim() !== '') {
                
                // Tentukan berapa multiplier yang ada
                const multiplierCount = this.countMultipliers(worksheet, R);
                
                // Buat formula berdasarkan jumlah multiplier
                let formula;
                if (multiplierCount === 2) {
                    // Kolom O = D * G
                    const cellG = XLSX.utils.encode_cell({ r: R, c: 6 });
                    formula = `${cellD}*${cellG}`;
                } else if (multiplierCount === 3) {
                    // Kolom O = D * G * J
                    const cellG = XLSX.utils.encode_cell({ r: R, c: 6 });
                    const cellJ = XLSX.utils.encode_cell({ r: R, c: 9 });
                    formula = `${cellD}*${cellG}*${cellJ}`;
                } else if (multiplierCount >= 4) {
                    // Untuk multiplier lebih dari 3, tambahkan semua
                    formula = this.buildMultiplierFormula(worksheet, R, multiplierCount);
                } else {
                    // Jika hanya 1 multiplier atau tidak ada, set ke nilai kolom D
                    formula = `${cellD}`;
                }
                
                // Set formula ke kolom O
                if (!worksheet[cellO]) {
                    worksheet[cellO] = {};
                }
                worksheet[cellO].f = formula;
                worksheet[cellO].t = 'n'; // Type numeric untuk formula
            }
        }
        
        return worksheet;
    }

    countMultipliers(worksheet, rowIndex) {
        let multiplierCount = 0;

        // Cek kolom D (index 3) - multiplier pertama
        const cellD = XLSX.utils.encode_cell({ r: rowIndex, c: 3 });
        if (worksheet[cellD] && worksheet[cellD].v && String(worksheet[cellD].v).trim() !== '') {
            multiplierCount++;
        }
        
        // Cek kolom G (index 6) - multiplier kedua
        const cellG = XLSX.utils.encode_cell({ r: rowIndex, c: 6 });
        if (worksheet[cellG] && worksheet[cellG].v && String(worksheet[cellG].v).trim() !== '') {
            multiplierCount++;
        }
        
        // Cek kolom J (index 9) - multiplier ketiga  
        const cellJ = XLSX.utils.encode_cell({ r: rowIndex, c: 9 });
        if (worksheet[cellJ] && worksheet[cellJ].v && String(worksheet[cellJ].v).trim() !== '') {
            multiplierCount++;
        }
        
        return multiplierCount;
    }

    buildMultiplierFormula(worksheet, rowIndex, multiplierCount) {
        const cellD = XLSX.utils.encode_cell({ r: rowIndex, c: 3 });
        let formula = cellD;
        
        // Multiplier columns: D(3), G(6), J(9), M(12), P(15), etc.
        const multiplierColumns = [3, 6, 9]; // Tambahkan lebih banyak jika diperlukan
        
        for (let i = 0; i < Math.min(multiplierCount, multiplierColumns.length); i++) {
            const multiplierCell = XLSX.utils.encode_cell({ r: rowIndex, c: multiplierColumns[i] });
            formula += `*${multiplierCell}`;
        }
        
        return formula;
    }

    addMultiplicationFormulasU(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Kolom U adalah index 20 (0-based)
        const targetColumn = 20;
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cellO = XLSX.utils.encode_cell({ r: R, c: 14 }); // Kolom O
            const cellS = XLSX.utils.encode_cell({ r: R, c: 18 }); // Kolom S
            const cellU = XLSX.utils.encode_cell({ r: R, c: targetColumn }); // Kolom U
            
            // Periksa apakah kolom O dan S tidak kosong
            if (worksheet[cellO] && worksheet[cellO].v && String(worksheet[cellO].v).trim() !== '' &&
                worksheet[cellS] && worksheet[cellS].v && String(worksheet[cellS].v).trim() !== '') {
                
                // Buat formula: U = O * S
                const formula = `${cellO}*${cellS}`;
                
                // Set formula ke kolom U
                if (!worksheet[cellU]) {
                    worksheet[cellU] = {};
                }
                worksheet[cellU].f = formula;
                worksheet[cellU].t = 'n'; // Type numeric untuk formula
            }
        }
        
        return worksheet;
    }

    addHierarchicalSumFormulas(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const formulasAdded = [];
        
        // Process from bottom to top for hierarchical sums
        for (let R = range.e.r; R >= range.s.r; --R) {
            const cellA = XLSX.utils.encode_cell({ r: R, c: 0 }); // Column A
            const cellB = XLSX.utils.encode_cell({ r: R, c: 1 }); // Column B
            const cellU = XLSX.utils.encode_cell({ r: R, c: 20 }); // Column U
            
            const valueA = worksheet[cellA] ? String(worksheet[cellA].v).trim() : '';
            const valueB = worksheet[cellB] ? String(worksheet[cellB].v).trim() : '';
            
            let formula = null;
            
            console.log(`Row ${R}: A="${valueA}", B="${valueB}"`); // Debug log
            
            // Rule 1: Column B has ">" - CHECK THIS FIRST
            if (valueB.includes('>')) {
                formula = this.createSumForGreaterThan(worksheet, R);
                console.log(`Rule 1 applied to row ${R}: ${formula}`); // Debug log
            }
            // Rule 2: Column A has 6-digit number
            else if (/^\d{6}$/.test(valueA)) {
                formula = this.createSumFor6Digit(worksheet, R);
            }
            // Rule 3: Column A has single alphabet
            else if (/^[A-Za-z]$/.test(valueA)) {
                formula = this.createSumForSingleAlphabet(worksheet, R);
            }
            // Rule 4: Column A has 3-digit number
            else if (/^\d{3}$/.test(valueA)) {
                formula = this.createSumFor3Digit(worksheet, R);
            }
            // Rule 5: Column A has code 433 (4 digit. 3 alphabet. 3 digit)
            else if (/^\d{4}\.[A-Za-z]{3}\.\d{3}$/.test(valueA)) {
                formula = this.createSumFor433Code(worksheet, R);
            }
            // Rule 6: Column A has code 43 (4 digit. 3 alphabet)
            else if (/^\d{4}\.[A-Za-z]{3}$/.test(valueA)) {
                formula = this.createSumFor43Code(worksheet, R);
            }
            // Rule 7: Column A has 4-digit number
            else if (/^\d{4}$/.test(valueA)) {
                formula = this.createSumFor4Digit(worksheet, R);
            }
            // Rule 8: Column A has code 322 (3 digit. 2 digit. 2 alphabet)
            else if (/^\d{3}\.\d{2}\.[A-Za-z]{2}$/.test(valueA)) {
                formula = this.createSumFor322Code(worksheet, R);
            }
            
            if (formula) {
                if (!worksheet[cellU]) {
                    worksheet[cellU] = {};
                }
                worksheet[cellU].f = formula;
                worksheet[cellU].t = 'n';
                formulasAdded.push({ row: R, formula: formula });
                console.log(`Formula added to U${R+1}: ${formula}`); // Debug log
            }
        }
        
        console.log(`Added ${formulasAdded.length} hierarchical sum formulas`);
        return worksheet;
    }

    // Rule 1: Column B has ">" - FIXED VERSION
    createSumForGreaterThan(worksheet, startRow) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const sumCells = [];
        
        console.log(`Checking Rule 1 for row ${startRow}`); // Debug log
        
        for (let R = startRow + 1; R <= range.e.r; ++R) {
            const cellB = XLSX.utils.encode_cell({ r: R, c: 1 });
            const cellU = XLSX.utils.encode_cell({ r: R, c: 20 });
            
            const valueB = worksheet[cellB] ? String(worksheet[cellB].v).trim() : '';
            
            console.log(`  Row ${R}: B="${valueB}"`); // Debug log
            
            if (valueB === '-') {
                sumCells.push(cellU);
                console.log(`    Added ${cellU} to sum`); // Debug log
            } else {
                console.log(`    Stopping at row ${R} - value is "${valueB}"`); // Debug log
                break; // Stop when we find a non-dash in column B
            }
        }
        
        const formula = sumCells.length > 0 ? `=SUM(${sumCells.join(',')})` : null;
        console.log(`Rule 1 result: ${formula}`); // Debug log
        return formula;
    }

    // Rule 2: 6-digit number in column A
    createSumFor6Digit(worksheet, startRow) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const sumCells = [];
        
        for (let R = startRow + 1; R <= range.e.r; ++R) {
            const cellA = XLSX.utils.encode_cell({ r: R, c: 0 });
            const cellB = XLSX.utils.encode_cell({ r: R, c: 1 });
            const cellU = XLSX.utils.encode_cell({ r: R, c: 20 });
            
            // Stop when we find next 6-digit number
            if (worksheet[cellA] && /^\d{6}$/.test(String(worksheet[cellA].v).trim())) {
                break;
            }
            
            // Include cells that have "-" in column B
            if (worksheet[cellB] && String(worksheet[cellB].v).trim() === '-') {
                sumCells.push(cellU);
            }
        }
        
        return sumCells.length > 0 ? `=SUM(${sumCells.join(',')})` : null;
    }

    // Rule 3: Single alphabet in column A
    createSumForSingleAlphabet(worksheet, startRow) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const sumCells = [];
        
        for (let R = startRow + 1; R <= range.e.r; ++R) {
            const cellA = XLSX.utils.encode_cell({ r: R, c: 0 });
            const cellU = XLSX.utils.encode_cell({ r: R, c: 20 });
            
            // Stop when we find next single alphabet
            if (worksheet[cellA] && /^[A-Za-z]$/.test(String(worksheet[cellA].v).trim())) {
                break;
            }
            
            // Include cells that have 6-digit numbers
            if (worksheet[cellA] && /^\d{6}$/.test(String(worksheet[cellA].v).trim())) {
                sumCells.push(cellU);
            }
        }
        
        return sumCells.length > 0 ? `=SUM(${sumCells.join(',')})` : null;
    }

    // Rule 4: 3-digit number in column A
    createSumFor3Digit(worksheet, startRow) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const sumCells = [];
        
        for (let R = startRow + 1; R <= range.e.r; ++R) {
            const cellA = XLSX.utils.encode_cell({ r: R, c: 0 });
            const cellU = XLSX.utils.encode_cell({ r: R, c: 20 });
            
            // Stop when we find next 3-digit number
            if (worksheet[cellA] && /^\d{3}$/.test(String(worksheet[cellA].v).trim())) {
                break;
            }
            
            // Include cells that have single alphabet
            if (worksheet[cellA] && /^[A-Za-z]$/.test(String(worksheet[cellA].v).trim())) {
                sumCells.push(cellU);
            }
        }
        
        return sumCells.length > 0 ? `=SUM(${sumCells.join(',')})` : null;
    }

    // Rule 5: Code 433 in column A
    createSumFor433Code(worksheet, startRow) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const sumCells = [];
        
        for (let R = startRow + 1; R <= range.e.r; ++R) {
            const cellA = XLSX.utils.encode_cell({ r: R, c: 0 });
            const cellU = XLSX.utils.encode_cell({ r: R, c: 20 });
            
            // Stop when we find next 433 code
            if (worksheet[cellA] && /^\d{4}\.[A-Za-z]{3}\.\d{3}$/.test(String(worksheet[cellA].v).trim())) {
                break;
            }
            
            // Include cells that have 3-digit numbers
            if (worksheet[cellA] && /^\d{3}$/.test(String(worksheet[cellA].v).trim())) {
                sumCells.push(cellU);
            }
        }
        
        return sumCells.length > 0 ? `=SUM(${sumCells.join(',')})` : null;
    }

    // Rule 6: Code 43 in column A
    createSumFor43Code(worksheet, startRow) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const sumCells = [];
        
        for (let R = startRow + 1; R <= range.e.r; ++R) {
            const cellA = XLSX.utils.encode_cell({ r: R, c: 0 });
            const cellU = XLSX.utils.encode_cell({ r: R, c: 20 });
            
            // Stop when we find next 43 code
            if (worksheet[cellA] && /^\d{4}\.[A-Za-z]{3}$/.test(String(worksheet[cellA].v).trim())) {
                break;
            }
            
            // Include cells that have 433 codes
            if (worksheet[cellA] && /^\d{4}\.[A-Za-z]{3}\.\d{3}$/.test(String(worksheet[cellA].v).trim())) {
                sumCells.push(cellU);
            }
        }
        
        return sumCells.length > 0 ? `=SUM(${sumCells.join(',')})` : null;
    }

    // Rule 7: 4-digit number in column A
    createSumFor4Digit(worksheet, startRow) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const sumCells = [];
        
        for (let R = startRow + 1; R <= range.e.r; ++R) {
            const cellA = XLSX.utils.encode_cell({ r: R, c: 0 });
            const cellU = XLSX.utils.encode_cell({ r: R, c: 20 });
            
            // Stop when we find next 4-digit number
            if (worksheet[cellA] && /^\d{4}$/.test(String(worksheet[cellA].v).trim())) {
                break;
            }
            
            // Include cells that have 43 codes
            if (worksheet[cellA] && /^\d{4}\.[A-Za-z]{3}$/.test(String(worksheet[cellA].v).trim())) {
                sumCells.push(cellU);
            }
        }
        
        return sumCells.length > 0 ? `=SUM(${sumCells.join(',')})` : null;
    }

    // Rule 8: Code 322 in column A
    createSumFor322Code(worksheet, startRow) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const sumCells = [];
        
        for (let R = startRow + 1; R <= range.e.r; ++R) {
            const cellA = XLSX.utils.encode_cell({ r: R, c: 0 });
            const cellU = XLSX.utils.encode_cell({ r: R, c: 20 });
            
            // Stop when we find next 322 code
            if (worksheet[cellA] && /^\d{3}\.\d{2}\.[A-Za-z]{2}$/.test(String(worksheet[cellA].v).trim())) {
                break;
            }
            
            // Include cells that have 4-digit numbers
            if (worksheet[cellA] && /^\d{4}$/.test(String(worksheet[cellA].v).trim())) {
                sumCells.push(cellU);
            }
        }
        
        return sumCells.length > 0 ? `=SUM(${sumCells.join(',')})` : null;
    }

    processFile() {
        if (!this.workbook) {
            this.updateStatus('❌ No file loaded', 'error');
            return;
        }

        this.updateStatus('🔄 Processing file: Unmerging cells, unwrapping text, deleting columns B & C, adding blank columns, text-to-columns, adding multiplication formulas (O and U), hierarchical sum formulas, and applying number formatting...', 'info');

        try {
            // Create a copy of the workbook
            this.processedWorkbook = XLSX.utils.book_new();

            // Process each worksheet
            this.workbook.SheetNames.forEach(sheetName => {
                let processedSheet = this.workbook.Sheets[sheetName];

                // Step 1: Unmerge all cells by removing merge ranges
                if (processedSheet['!merges']) {
                    processedSheet['!merges'] = [];
                }

                // Step 2: Process each cell to unwrap text
                const range = XLSX.utils.decode_range(processedSheet['!ref']);
                
                for (let R = range.s.r; R <= range.e.r; ++R) {
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                        
                        if (processedSheet[cellAddress]) {
                            // Unwrap text (remove formatting and keep plain text)
                            processedSheet[cellAddress].v = this.unwrapText(processedSheet[cellAddress].v);
                            
                            // Force string type to remove any formatting
                            processedSheet[cellAddress].t = 's';
                        }
                    }
                }

                // Step 3: Delete columns B and C (index 1 and 2)
                processedSheet = this.deleteColumns(processedSheet, [1, 2]);
                
                // Step 4: Add 3 blank columns after column E (index 4)
                processedSheet = this.addBlankColumns(processedSheet, 4, 3);
                
                // Step 5: Perform text-to-columns on column E (index 4) using space delimiter
                processedSheet = this.textToColumns(processedSheet, 4, ' ');

                // Step 6: Add 1 blank columns after column E (index 4)
                processedSheet = this.addBlankColumns(processedSheet, 4, 1);
                
                // Step 7: Perform text-to-columns on column E (index 4) using space delimiter
                processedSheet = this.textToColumns(processedSheet, 4, '.');

                // Step 8: Delete column F (index 5)
                processedSheet = this.deleteColumns(processedSheet, [5]);

                // Step 9: Perform text-to-columns on column C (index 2) using [ delimiter
                processedSheet = this.textToColumns(processedSheet, 2, '[');

                // Step 10: Add 10 blank columns after column D (index 3)
                processedSheet = this.addBlankColumns(processedSheet, 3, 10);

                // Step 11: Perform text-to-columns on column D (index 3) using ] delimiter
                processedSheet = this.textToColumns(processedSheet, 3, ']');

                // Step 12: Perform text-to-columns on column D (index 3) using space delimiter
                processedSheet = this.textToColumns(processedSheet, 3, ' ');

                // Step 13: Add multiplication formulas in column O (index 14)
                processedSheet = this.addMultiplicationFormulas(processedSheet);

                // Step 14: Add multiplication formulas in column U (index 20) - O * S
                processedSheet = this.addMultiplicationFormulasU(processedSheet);

                // Step 15: Add hierarchical sum formulas in column U
                processedSheet = this.addHierarchicalSumFormulas(processedSheet);

                // Step 16: Apply number formatting to columns S and U
                processedSheet = this.applyNumberFormattingAsString(processedSheet);
                
                XLSX.utils.book_append_sheet(this.processedWorkbook, processedSheet, sheetName);
            });

            // Show download button
            this.downloadBtn.classList.remove('hidden');
            this.downloadBtn.classList.add('btn-success');
            
            this.updateStatus('✅ File processed successfully! Click download button below.', 'success');
            this.dropZone.innerHTML = '<h3>✅ Processing Complete</h3><p>Your file is ready for download</p>';
            
        } catch (error) {
            this.updateStatus('❌ Error processing file: ' + error.message, 'error');
            this.resetUploadArea();
        }
    }

    applyNumberFormattingAsString(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        const columnsToFormat = [18, 20]; // S and U
        
        columnsToFormat.forEach(columnIndex => {
            for (let R = range.s.r; R <= range.e.r; ++R) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: columnIndex });
                
                if (worksheet[cellAddress]) {
                    const cell = worksheet[cellAddress];
                    
                    if (cell.v !== null && cell.v !== undefined && !cell.f) {
                        const cellValue = cell.v;
                        let numericValue;
                        
                        if (typeof cellValue === 'number') {
                            numericValue = cellValue;
                        } else if (typeof cellValue === 'string') {
                            const cleanValue = cellValue.replace(/\./g, '').replace(',', '.');
                            numericValue = parseFloat(cleanValue);
                        } else {
                            continue;
                        }
                        
                        if (!isNaN(numericValue) && isFinite(numericValue)) {
                            // Format as string with Indonesian thousands separators
                            const formattedValue = this.formatNumberWithSeparators(numericValue);
                            worksheet[cellAddress].v = formattedValue;
                            worksheet[cellAddress].t = 's'; // Force string type
                        }
                    }
                }
            }
        });
        
        return worksheet;
    }

    formatNumberWithSeparators(number) {
        return number.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    }

    addBlankColumns(worksheet, afterColumnIndex, numColumns) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Expand the range to accommodate new columns
        range.e.c += numColumns;
        
        // Shift existing columns to the right to make space
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.e.c; C > afterColumnIndex; --C) {
                const currentCell = XLSX.utils.encode_cell({ r: R, c: C });
                const sourceCell = XLSX.utils.encode_cell({ r: R, c: C - numColumns });
                
                if (worksheet[sourceCell]) {
                    worksheet[currentCell] = worksheet[sourceCell];
                } else {
                    delete worksheet[currentCell];
                }
            }
            
            // Clear the new blank columns
            for (let C = afterColumnIndex + 1; C <= afterColumnIndex + numColumns; ++C) {
                const blankCell = XLSX.utils.encode_cell({ r: R, c: C });
                delete worksheet[blankCell]; // Ensure they're empty
            }
        }
        
        // Update the worksheet range
        worksheet['!ref'] = XLSX.utils.encode_range(range);
        
        return worksheet;
    }

    textToColumns(worksheet, sourceColumnIndex, delimiter = ',') {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const sourceCell = XLSX.utils.encode_cell({ r: R, c: sourceColumnIndex });
            
            if (worksheet[sourceCell] && worksheet[sourceCell].v) {
                const cellValue = String(worksheet[sourceCell].v);
                
                // Split the text by delimiter
                const splitValues = cellValue.split(delimiter).map(val => val.trim());
                
                // Write split values to consecutive columns starting from source column
                splitValues.forEach((value, index) => {
                    const targetCell = XLSX.utils.encode_cell({ r: R, c: sourceColumnIndex + index });
                    
                    if (!worksheet[targetCell]) {
                        worksheet[targetCell] = {};
                    }
                    
                    worksheet[targetCell].v = value;
                    worksheet[targetCell].t = 's'; // Force string type
                });
                
                // Clear remaining cells in the row if split values are fewer than previous splits
                for (let C = sourceColumnIndex + splitValues.length; C <= range.e.c; ++C) {
                    const remainingCell = XLSX.utils.encode_cell({ r: R, c: C });
                    // Don't delete if it's beyond our source column data
                    if (C > sourceColumnIndex + splitValues.length - 1) {
                        // Keep existing data in other columns
                        continue;
                    }
                }
            }
        }
        
        return worksheet;
    }

    deleteColumns(worksheet, columnsToDelete) {
        // Sort columns in descending order to avoid index issues when deleting
        columnsToDelete.sort((a, b) => b - a);
        
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Process each row
        for (let R = range.s.r; R <= range.e.r; ++R) {
            // Delete specified columns for this row
            columnsToDelete.forEach(colIndex => {
                // Shift all cells to the right of the deleted column to the left
                for (let C = colIndex; C <= range.e.c; ++C) {
                    const currentCell = XLSX.utils.encode_cell({ r: R, c: C });
                    const nextCell = XLSX.utils.encode_cell({ r: R, c: C + 1 });
                    
                    if (worksheet[nextCell]) {
                        // Move next cell to current position
                        worksheet[currentCell] = worksheet[nextCell];
                    } else {
                        // If no next cell, delete current cell
                        delete worksheet[currentCell];
                    }
                }
                
                // Delete the last column in the row since we shifted everything left
                const lastCell = XLSX.utils.encode_cell({ r: R, c: range.e.c });
                delete worksheet[lastCell];
            });
        }
        
        // Update the range to reflect the deleted columns
        range.e.c -= columnsToDelete.length;
        worksheet['!ref'] = XLSX.utils.encode_range(range);
        
        // Update merge ranges if they exist (adjust column indexes)
        if (worksheet['!merges']) {
            worksheet['!merges'] = worksheet['!merges'].map(merge => {
                const newMerge = {...merge};
                
                // Adjust start column
                if (newMerge.s.c >= columnsToDelete[0]) {
                    newMerge.s.c -= columnsToDelete.filter(col => col <= newMerge.s.c).length;
                }
                
                // Adjust end column  
                if (newMerge.e.c >= columnsToDelete[0]) {
                    newMerge.e.c -= columnsToDelete.filter(col => col <= newMerge.e.c).length;
                }
                
                return newMerge;
            }).filter(merge => merge.s.c <= range.e.c && merge.e.c <= range.e.c);
        }

        return worksheet;
    }

    unwrapText(value) {
        if (value === null || value === undefined) return '';
        
        // Convert to string
        const text = String(value);
        
        // Remove extra whitespace, newlines, and normalize spaces
        return text
            .replace(/\r\n/g, ' ')  // Windows newlines
            .replace(/\n/g, ' ')    // Unix newlines  
            .replace(/\t/g, ' ')    // Tabs
            .replace(/\s+/g, ' ')   // Multiple spaces to single space
            .trim();                // Trim edges
    }

    downloadFile() {
        if (!this.processedWorkbook) {
            this.updateStatus('❌ No processed file available', 'error');
            return;
        }

        try {
            // Generate filename with timestamp
            const originalName = this.fileInput.files[0]?.name.replace(/\.[^/.]+$/, "") || 'processed';
            const timestamp = new Date().toISOString().slice(0, 19).replace(/[:]/g, '-');
            const filename = `${originalName}_processed_${timestamp}.xlsx`;
            
            // Download the file
            XLSX.writeFile(this.processedWorkbook, filename);
            this.updateStatus(`✅ File downloaded: ${filename}`, 'success');
            
        } catch (error) {
            this.updateStatus('❌ Error downloading file: ' + error.message, 'error');
        }
    }

    resetUploadArea() {
        this.dropZone.innerHTML = '<h3>📁 Upload Excel File</h3><p>Click here or drag & drop your Excel file</p>';
        this.dropZone.style.borderColor = '#007cba';
        this.dropZone.style.backgroundColor = '#f8f9fa';
    }

    updateStatus(message, type) {
        this.status.textContent = message;
        this.status.className = `status ${type}`;
        console.log(`[${type}] ${message}`);
    }
}

// Initialize the application when page loads
document.addEventListener('DOMContentLoaded', () => {
    new ExcelProcessor();
});