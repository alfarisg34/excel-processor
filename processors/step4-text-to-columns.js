class Step4TextToColumns extends BaseProcessor {
    constructor() {
        super();
        this.sourceColumnIndex = 0; // Column A
        this.delimiter = ',';
        this.numBlankColumns = 10;
        this.insertAfterColumn = 4; // After column E
    }
    
    getName() {
        return "Convert text to columns";
    }
    
    process(worksheet) {
        if (!this.validateInput(worksheet)) {
            throw new Error('Invalid input for Step4TextToColumns');
        }
        
        // Add blank columns first
        worksheet = this.addBlankColumns(worksheet);
        
        // Perform text-to-columns
        worksheet = this.splitTextToColumns(worksheet);
        
        if (!this.validateOutput(worksheet)) {
            throw new Error('Invalid output from Step4TextToColumns');
        }
        
        return worksheet;
    }
    
    addBlankColumns(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        range.e.c += this.numBlankColumns;
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.e.c; C > this.insertAfterColumn; --C) {
                const currentCell = XLSX.utils.encode_cell({ r: R, c: C });
                const sourceCell = XLSX.utils.encode_cell({ r: R, c: C - this.numBlankColumns });
                
                if (worksheet[sourceCell]) {
                    worksheet[currentCell] = worksheet[sourceCell];
                } else {
                    delete worksheet[currentCell];
                }
            }
            
            for (let C = this.insertAfterColumn + 1; C <= this.insertAfterColumn + this.numBlankColumns; ++C) {
                const blankCell = XLSX.utils.encode_cell({ r: R, c: C });
                delete worksheet[blankCell];
            }
        }
        
        worksheet['!ref'] = XLSX.utils.encode_range(range);
        return worksheet;
    }
    
    splitTextToColumns(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const sourceCell = XLSX.utils.encode_cell({ r: R, c: this.sourceColumnIndex });
            
            if (worksheet[sourceCell] && worksheet[sourceCell].v) {
                const cellValue = String(worksheet[sourceCell].v);
                const splitValues = cellValue.split(this.delimiter).map(val => val.trim());
                
                splitValues.forEach((value, index) => {
                    const targetCell = XLSX.utils.encode_cell({ r: R, c: this.sourceColumnIndex + index });
                    
                    if (!worksheet[targetCell]) {
                        worksheet[targetCell] = {};
                    }
                    
                    worksheet[targetCell].v = value;
                    worksheet[targetCell].t = 's';
                });
            }
        }
        
        return worksheet;
    }
}