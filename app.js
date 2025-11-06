class ExcelProcessor {
    constructor() {
        this.workbook = null;
        this.processedWorkbook = null;
        this.init();
    }

    init() {
        // DOM elements
        this.fileInput = document.getElementById('fileInput');
        this.fileInputSemulaMenjadi = document.getElementById('fileInputSemulaMenjadi');
        this.dropZone = document.getElementById('dropZone');
        this.dropZoneSemulaMenjadi = document.getElementById('dropZoneSemulaMenjadi');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.status = document.getElementById('status');

        // Event listeners for standard processing
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        this.dropZone.addEventListener('click', () => this.fileInput.click());
        this.dropZone.addEventListener('dragover', (e) => this.handleDragOver(e));
        this.dropZone.addEventListener('drop', (e) => this.handleFileDrop(e));

        // Event listeners for semula menjadi processing
        this.fileInputSemulaMenjadi.addEventListener('change', (e) => this.handleFileSelectSemulaMenjadi(e));
        this.dropZoneSemulaMenjadi.addEventListener('click', () => this.fileInputSemulaMenjadi.click());
        this.dropZoneSemulaMenjadi.addEventListener('dragover', (e) => this.handleDragOverSemulaMenjadi(e));
        this.dropZoneSemulaMenjadi.addEventListener('drop', (e) => this.handleFileDropSemulaMenjadi(e));

        this.downloadBtn.addEventListener('click', () => this.downloadFile());

        this.updateStatus('Please upload an Excel file to begin processing', 'info');
    }

    processFile() {
        if (!this.workbook) {
            this.updateStatus('❌ No file loaded', 'error');
            return;
        }

        this.updateStatus('🔄 Processing file: Unmerging cells, unwrapping text, clearing blank cells, inserting rows around Jakarta, deleting columns B & C, adding blank columns, text-to-columns, adding multiplication formulas (O and U), hierarchical sum formulas, applying number formatting, auto-fitting columns, and moving column Y to W...', 'info');

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
                        }
                    }
                }

                // Step 3: Clear blank cells to fix text flow
                processedSheet = this.clearBlankCells(processedSheet);

                // NEW STEP: Insert rows above and below cells containing "Jakarta"
                processedSheet = this.insertRowsAroundJakarta(processedSheet);

                // Step 4: Delete columns B and C (index 1 and 2)
                processedSheet = this.deleteColumns(processedSheet, [1, 2]);
                
                // Step 5: Add 3 blank columns after column E (index 4)
                processedSheet = this.addBlankColumns(processedSheet, 4, 3);
                
                // Step 6: Perform text-to-columns on column E (index 4) using space delimiter
                processedSheet = this.textToColumns(processedSheet, 4, ' ');

                // Step 7: Add 1 blank columns after column E (index 4)
                processedSheet = this.addBlankColumns(processedSheet, 4, 1);
                
                // Step 8: Perform text-to-columns on column E (index 4) using space delimiter
                processedSheet = this.textToColumns(processedSheet, 4, '.');

                // Step 9: Delete column F (index 5)
                processedSheet = this.deleteColumns(processedSheet, [5]);

                // Step 10: Perform text-to-columns on column C (index 2) using [ delimiter
                processedSheet = this.textToColumns(processedSheet, 2, '[');

                // Step 11: Add 10 blank columns after column D (index 3)
                processedSheet = this.addBlankColumns(processedSheet, 3, 10);

                // Step 12: Perform text-to-columns on column D (index 3) using ] delimiter
                processedSheet = this.textToColumns(processedSheet, 3, ']');

                // Step 13: Perform text-to-columns on column D (index 3) using space delimiter
                processedSheet = this.textToColumns(processedSheet, 3, ' ');

                // Step 14: Add multiplication formulas in column O (index 14)
                processedSheet = this.addMultiplicationFormulas(processedSheet);

                // Step 15: Add multiplication formulas in column U (index 20) - O * S
                processedSheet = this.addMultiplicationFormulasU(processedSheet);

                // NEW STEP: Convert column S values to numbers
                processedSheet = this.convertColumnSToNumbers(processedSheet);

                // Step 16: Add hierarchical sum formulas in column U
                processedSheet = this.addHierarchicalSumFormulas(processedSheet);

                // Step 17: Apply number formatting to columns S and U
                processedSheet = this.applyNumberFormattingAsString(processedSheet);

                // Step 18: Auto-fit columns A and D-W
                processedSheet = this.autoFitColumns(processedSheet);

                // Step 19: Move text from column Y to column W
                processedSheet = this.moveColumnYToW(processedSheet);
                
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

    processFileSemulaMenjadi() {
        if (!this.workbook) {
            this.updateStatus('❌ No file loaded', 'error');
            return;
        }

        this.updateStatus('🔄 Processing file with Semula Menjadi: First processing file, then applying additional steps (adding 3 rows at top)...', 'info');

        try {
            // First, do the normal processFile()
            this.processFile();
            
            // Wait a bit for the first processing to complete, then do additional steps
            setTimeout(() => {
                this.updateStatus('🔄 Applying additional Semula Menjadi steps (adding 3 rows at top)...', 'info');
                
                // Process each worksheet for additional steps
                this.workbook.SheetNames.forEach(sheetName => {
                    let processedSheet = this.processedWorkbook.Sheets[sheetName];
                    
                    // STEP 1: Add 3 rows at the top of the sheet
                    processedSheet = this.addThreeRowsAtTop(processedSheet);
                    
                    // Update the processed sheet
                    this.processedWorkbook.Sheets[sheetName] = processedSheet;
                });
                
                this.updateStatus('✅ Semula Menjadi processing complete! Click download button below.', 'success');
                
            }, 1000);
            
        } catch (error) {
            this.updateStatus('❌ Error in Semula Menjadi processing: ' + error.message, 'error');
            this.resetUploadArea();
        }
    }

    // Add this new method to insert 3 rows at the top
    addThreeRowsAtTop(worksheet) {
        console.log('Adding 3 rows at the top of the sheet...');
        
        // Insert 3 rows at the top (row 0)
        for (let i = 0; i < 3; i++) {
            worksheet = this.insertRowAtTop(worksheet);
        }
        
        console.log('Added 3 rows at the top of the sheet');
        return worksheet;
    }

    // Helper method to insert a single row at the top
    insertRowAtTop(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Increase the row range by 1
        range.e.r += 1;
        
        // Shift all existing rows down by 1
        for (let R = range.e.r; R > 0; --R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const currentCell = XLSX.utils.encode_cell({ r: R, c: C });
                const sourceCell = XLSX.utils.encode_cell({ r: R - 1, c: C });
                
                if (worksheet[sourceCell]) {
                    worksheet[currentCell] = { ...worksheet[sourceCell] };
                    
                    // Update cell references in formulas if they exist
                    if (worksheet[currentCell].f) {
                        worksheet[currentCell].f = this.updateFormulaReferences(worksheet[currentCell].f, 1);
                    }
                } else {
                    delete worksheet[currentCell];
                }
            }
        }
        
        // Clear the new top row (row 0)
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const newCell = XLSX.utils.encode_cell({ r: 0, c: C });
            delete worksheet[newCell];
        }
        
        // Update merge ranges if they exist
        if (worksheet['!merges']) {
            worksheet['!merges'] = worksheet['!merges'].map(merge => {
                const newMerge = { ...merge };
                
                // Adjust all row indices since we're inserting at the top
                newMerge.s.r += 1;
                newMerge.e.r += 1;
                
                return newMerge;
            });
        }
        
        // Update the worksheet range
        worksheet['!ref'] = XLSX.utils.encode_range(range);
        
        return worksheet;
    }

    // Helper method to update formula references when rows are inserted
    updateFormulaReferences(formula, rowOffset) {
        if (!formula) return formula;
        
        // Simple regex to match cell references like A1, B2, etc.
        const cellRefRegex = /([A-Z]+)(\d+)/g;
        
        return formula.replace(cellRefRegex, (match, col, row) => {
            const rowNum = parseInt(row);
            return `${col}${rowNum + rowOffset}`;
        });
    }

    convertColumnSToNumbers(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Column S is index 18 (0-based)
        const columnSIndex = 18;
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: columnSIndex });
            
            if (worksheet[cellAddress]) {
                const cell = worksheet[cellAddress];
                const originalValue = cell.v;
                
                // Skip if cell already has a formula
                if (cell.f) {
                    continue;
                }
                
                // Convert to number if it's a string that represents a number
                if (typeof originalValue === 'string') {
                    // Remove any formatting characters (thousands separators, currency symbols, etc.)
                    const cleanValue = originalValue
                        .replace(/\./g, '')  // Remove thousands separators (dots)
                        .replace(/,/g, '.')  // Convert comma decimal to dot decimal
                        .replace(/[^\d.-]/g, '') // Remove any non-numeric characters except dot and minus
                        .trim();
                    
                    // Convert to number
                    const numericValue = parseFloat(cleanValue);
                    
                    // If it's a valid number, update the cell
                    if (!isNaN(numericValue) && isFinite(numericValue)) {
                        worksheet[cellAddress].v = numericValue;
                        worksheet[cellAddress].t = 'n'; // Set type to number
                        
                        console.log(`Converted S${R + 1}: "${originalValue}" -> ${numericValue}`);
                    } else if (cleanValue === '' || originalValue.trim() === '') {
                        // If empty string, set to empty
                        worksheet[cellAddress].v = '';
                        worksheet[cellAddress].t = 's'; // Set type to string
                    }
                    // If not a valid number, leave as is (it might be text)
                } else if (typeof originalValue === 'number') {
                    // Already a number, ensure type is set correctly
                    worksheet[cellAddress].t = 'n';
                }
            }
        }
        
        console.log('Converted column S values to numbers');
        return worksheet;
    }

    handleDragOverSemulaMenjadi(e) {
        e.preventDefault();
        this.dropZoneSemulaMenjadi.style.borderColor = '#28a745';
        this.dropZoneSemulaMenjadi.style.backgroundColor = '#e8f5e8';
    }

    handleFileDropSemulaMenjadi(e) {
        e.preventDefault();
        this.dropZoneSemulaMenjadi.style.borderColor = '#28a745';
        this.dropZoneSemulaMenjadi.style.backgroundColor = '#f8f9fa';
        
        const files = e.dataTransfer.files;
        if (files.length) {
            this.loadFileSemulaMenjadi(files[0]);
        }
    }

    handleFileSelectSemulaMenjadi(e) {
        const files = e.target.files;
        if (files.length) {
            this.loadFileSemulaMenjadi(files[0]);
        }
    }

    loadFileSemulaMenjadi(file) {
        if (!file.name.match(/\.(xlsx|xls|csv)$/)) {
            this.updateStatus('❌ Please upload a valid Excel file (.xlsx, .xls, .csv)', 'error');
            return;
        }

        this.updateStatus('⏳ Loading file for Semula Menjadi processing...', 'info');
        this.dropZoneSemulaMenjadi.innerHTML = '<h3>⏳ Processing...</h3><p>Semula Menjadi - Please wait</p>';

        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                this.workbook = XLSX.read(data, { type: 'array' });
                this.updateStatus('✅ File loaded! Starting Semula Menjadi processing...', 'info');
                
                // Automatically start Semula Menjadi processing
                setTimeout(() => this.processFileSemulaMenjadi(), 500);
                
            } catch (error) {
                this.updateStatus('❌ Error reading file: ' + error.message, 'error');
                this.resetUploadAreaSemulaMenjadi();
            }
        };

        reader.onerror = () => {
            this.updateStatus('❌ Error reading file', 'error');
            this.resetUploadAreaSemulaMenjadi();
        };

        reader.readAsArrayBuffer(file);
    }

    resetUploadAreaSemulaMenjadi() {
        this.dropZoneSemulaMenjadi.innerHTML = '<h3>🚀 Upload for Semula Menjadi</h3><p>Click here or drag & drop your Excel file</p>';
        this.dropZoneSemulaMenjadi.style.borderColor = '#28a745';
        this.dropZoneSemulaMenjadi.style.backgroundColor = '#f8f9fa';
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

    insertRowsAroundJakarta(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        let rowsToInsert = [];
        
        // First, find all rows that contain "Jakarta" in any cell
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                
                if (worksheet[cellAddress] && worksheet[cellAddress].v) {
                    const cellValue = String(worksheet[cellAddress].v);
                    
                    if (cellValue.toLowerCase().includes('jakarta,')) {
                        // Found a row with "Jakarta", mark it for row insertion
                        if (!rowsToInsert.includes(R)) {
                            rowsToInsert.push(R);
                        }
                        break; // No need to check other cells in this row
                    }
                }
            }
        }
        
        // Sort in descending order to maintain correct indices when inserting
        rowsToInsert.sort((a, b) => b - a);
        
        console.log(`Found ${rowsToInsert.length} rows containing "Jakarta"`);
        
        // Insert rows above and below each Jakarta row
        rowsToInsert.forEach(rowIndex => {
            console.log(`Inserting rows around row ${rowIndex + 1}`);
            
            // Insert row BELOW first (so it doesn't affect the row indices for ABOVE insertion)
            this.insertRow(worksheet, rowIndex + 1); // Insert below current row
            
            // Insert row ABOVE
            this.insertRow(worksheet, rowIndex); // Insert above current row (now rowIndex is still correct)
        });
        
        // Update the worksheet range after all insertions
        this.updateWorksheetRange(worksheet);
        
        console.log(`Inserted ${rowsToInsert.length * 2} rows around Jakarta cells`);
        return worksheet;
    }

    insertRow(worksheet, atRow) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Increase the row range
        range.e.r += 1;
        
        // Shift all rows from the insertion point downward
        for (let R = range.e.r; R > atRow; --R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const currentCell = XLSX.utils.encode_cell({ r: R, c: C });
                const sourceCell = XLSX.utils.encode_cell({ r: R - 1, c: C });
                
                if (worksheet[sourceCell]) {
                    worksheet[currentCell] = { ...worksheet[sourceCell] };
                } else {
                    delete worksheet[currentCell];
                }
            }
        }
        
        // Clear the newly inserted row
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const newCell = XLSX.utils.encode_cell({ r: atRow, c: C });
            delete worksheet[newCell];
        }
        
        // Update merge ranges if they exist
        if (worksheet['!merges']) {
            worksheet['!merges'] = worksheet['!merges'].map(merge => {
                const newMerge = { ...merge };
                
                // Adjust row indices for merges below the insertion point
                if (newMerge.s.r >= atRow) {
                    newMerge.s.r += 1;
                }
                if (newMerge.e.r >= atRow) {
                    newMerge.e.r += 1;
                }
                
                return newMerge;
            });
        }
    }

    moveColumnYToW(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Column Y is index 24 (0-based), Column W is index 22 (0-based)
        const sourceColumn = 24; // Y
        const targetColumn = 22; // W
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const sourceCell = XLSX.utils.encode_cell({ r: R, c: sourceColumn });
            const targetCell = XLSX.utils.encode_cell({ r: R, c: targetColumn });
            
            // Check if source cell has text content
            if (worksheet[sourceCell] && worksheet[sourceCell].v && 
                String(worksheet[sourceCell].v).trim() !== '') {
                
                const sourceValue = worksheet[sourceCell].v;
                
                // Move the value to column W
                if (!worksheet[targetCell]) {
                    worksheet[targetCell] = {};
                }
                worksheet[targetCell].v = sourceValue;
                
                // Copy any formatting if it exists
                if (worksheet[sourceCell].s) {
                    worksheet[targetCell].s = { ...worksheet[sourceCell].s };
                }
                
                // Copy cell type if specified
                if (worksheet[sourceCell].t) {
                    worksheet[targetCell].t = worksheet[sourceCell].t;
                }
                
                // Clear the source cell after moving
                delete worksheet[sourceCell];
            }
        }
        
        console.log('Moved text from column Y to column W');
        return worksheet;
    }

    unwrapText(value) {
        if (value === null || value === undefined) return '';
        
        // Convert to string and clean
        const text = String(value);
        
        // Remove extra whitespace, newlines, and normalize spaces
        const cleanedText = text
            .replace(/\r\n/g, ' ')  // Windows newlines
            .replace(/\n/g, ' ')    // Unix newlines  
            .replace(/\t/g, ' ')    // Tabs
            .replace(/\s+/g, ' ')   // Multiple spaces to single space
            .trim();                // Trim edges
        
        // Return empty string if result is just whitespace
        return cleanedText === '' ? '' : cleanedText;
    }

    clearBlankCells(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                
                if (worksheet[cellAddress]) {
                    const cell = worksheet[cellAddress];
                    
                    // Check if cell is effectively blank
                    if (cell.v === null || cell.v === undefined || 
                        (typeof cell.v === 'string' && cell.v.trim() === '') ||
                        cell.v === '') {
                        
                        // Completely remove the cell to allow text flow
                        delete worksheet[cellAddress];
                    } else {
                        // For non-blank cells, remove forced string type to allow text flow
                        // Only set type if it's actually a string with content
                        if (typeof cell.v === 'string' && cell.v.trim() !== '') {
                            // Don't force 's' type - let Excel handle the type automatically
                            delete cell.t; // Remove forced type
                        }
                    }
                }
            }
        }
        
        // Recalculate the range after cleaning
        this.updateWorksheetRange(worksheet);
        
        console.log('Cleared blank cells and fixed text flow');
        return worksheet;
    }

    updateWorksheetRange(worksheet) {
        // Find the actual used range
        let minRow = Infinity, maxRow = -Infinity, minCol = Infinity, maxCol = -Infinity;
        let hasData = false;
        
        for (const cellAddress in worksheet) {
            if (cellAddress[0] === '!') continue; // Skip special properties
            
            const cellRef = XLSX.utils.decode_cell(cellAddress);
            minRow = Math.min(minRow, cellRef.r);
            maxRow = Math.max(maxRow, cellRef.r);
            minCol = Math.min(minCol, cellRef.c);
            maxCol = Math.max(maxCol, cellRef.c);
            hasData = true;
        }
        
        if (hasData) {
            worksheet['!ref'] = XLSX.utils.encode_range({
                s: { r: minRow, c: minCol },
                e: { r: maxRow, c: maxCol }
            });
        } else {
            // If no data, set a minimal range
            worksheet['!ref'] = 'A1:A1';
        }
    }

    autoFitColumns(worksheet) {
        // Set specific widths for columns A and D-W
        const columnWidths = {
            0: 12,  // A - Wider for codes and numbers
            3: 4.5,  // D
            4: 4.5,  // E  
            5: 1.17,  // F
            6: 4.5,  // G
            7: 4.5,  // H
            8: 1.17,  // I
            9: 4.5,  // J
            10: 4.5, // K
            11: 1.17, // L
            12: 4.5, // M
            13: 4.5, // N
            14: 5.2, // O - Wider for formulas
            15: 11.2, // P
            16: 12, // Q
            17: 12, // R
            18: 11.75, // S - Wider for formatted numbers
            19: 12, // T
            20: 15, // U - Wider for formulas and formatted numbers
            21: 9.1, // V
            22: 3.5  // W
        };
        
        if (!worksheet['!cols']) {
            worksheet['!cols'] = [];
        }
        
        Object.keys(columnWidths).forEach(col => {
            const colIndex = parseInt(col);
            const width = columnWidths[colIndex];
            
            worksheet['!cols'][colIndex] = {
                wch: width,
                width: width
            };
        });
        
        console.log('Applied fixed column widths to A and D-W');
        return worksheet;
    }

    applyNumberFormattingAsString(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        const columnsToFormat = [20]; // U
        
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
                    // DON'T force string type - let Excel handle it automatically
                    // This allows text to flow into adjacent cells
                });
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
            const date = new Date();
            date.setHours(date.getHours() + 7); // add +7 hours
            const timestamp = date.toISOString().slice(0, 19).replace(/[:]/g, '-');
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