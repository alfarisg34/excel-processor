class Step2Unwrap extends BaseProcessor {
    getName() {
        return "Unwrap all text";
    }
    
    process(worksheet) {
        if (!this.validateInput(worksheet)) {
            throw new Error('Invalid input for Step2Unwrap');
        }
        
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                
                if (worksheet[cellAddress]) {
                    worksheet[cellAddress].v = this.unwrapText(worksheet[cellAddress].v);
                    worksheet[cellAddress].t = 's';
                }
            }
        }
        
        if (!this.validateOutput(worksheet)) {
            throw new Error('Invalid output from Step2Unwrap');
        }
        
        return worksheet;
    }
    
    unwrapText(value) {
        if (value === null || value === undefined) return '';
        const text = String(value);
        return text
            .replace(/\r\n/g, ' ')
            .replace(/\n/g, ' ')  
            .replace(/\t/g, ' ') 
            .replace(/\s+/g, ' ')
            .trim();
    }
}