class Step3DeleteColumns extends BaseProcessor {
    constructor() {
        super();
        this.columnsToDelete = [1, 2]; // Columns B and C
    }
    
    getName() {
        return "Delete columns B and C";
    }
    
    process(worksheet) {
        if (!this.validateInput(worksheet)) {
            throw new Error('Invalid input for Step3DeleteColumns');
        }
        
        // Sort columns in descending order to avoid index issues
        this.columnsToDelete.sort((a, b) => b - a);
        
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Process each row
        for (let R = range.s.r; R <= range.e.r; ++R) {
            this.columnsToDelete.forEach(colIndex => {
                for (let C = colIndex; C <= range.e.c; ++C) {
                    const currentCell = XLSX.utils.encode_cell({ r: R, c: C });
                    const nextCell = XLSX.utils.encode_cell({ r: R, c: C + 1 });
                    
                    if (worksheet[nextCell]) {
                        worksheet[currentCell] = worksheet[nextCell];
                    } else {
                        delete worksheet[currentCell];
                    }
                }
                
                const lastCell = XLSX.utils.encode_cell({ r: R, c: range.e.c });
                delete worksheet[lastCell];
            });
        }
        
        // Update range
        range.e.c -= this.columnsToDelete.length;
        worksheet['!ref'] = XLSX.utils.encode_range(range);
        
        // Update merge ranges
        if (worksheet['!merges']) {
            worksheet['!merges'] = worksheet['!merges'].map(merge => {
                const newMerge = {...merge};
                
                if (newMerge.s.c >= this.columnsToDelete[0]) {
                    newMerge.s.c -= this.columnsToDelete.filter(col => col <= newMerge.s.c).length;
                }
                
                if (newMerge.e.c >= this.columnsToDelete[0]) {
                    newMerge.e.c -= this.columnsToDelete.filter(col => col <= newMerge.e.c).length;
                }
                
                return newMerge;
            }).filter(merge => merge.s.c <= range.e.c && merge.e.c <= range.e.c);
        }
        
        if (!this.validateOutput(worksheet)) {
            throw new Error('Invalid output from Step3DeleteColumns');
        }
        
        return worksheet;
    }
}