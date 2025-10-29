const Validators = {
    validateWorksheet(worksheet) {
        if (!worksheet) {
            console.error('Worksheet is null or undefined');
            return false;
        }
        
        if (!worksheet['!ref']) {
            console.error('Worksheet missing range reference (!ref)');
            return false;
        }
        
        try {
            XLSX.utils.decode_range(worksheet['!ref']);
        } catch (error) {
            console.error('Invalid worksheet range:', error);
            return false;
        }
        
        return true;
    },
    
    validateCellAddress(address) {
        if (typeof address !== 'string') return false;
        return /^[A-Z]+[1-9]\d*$/.test(address);
    },
    
    validateColumnIndex(worksheet, columnIndex) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        return columnIndex >= range.s.c && columnIndex <= range.e.c;
    },
    
    validateRowIndex(worksheet, rowIndex) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        return rowIndex >= range.s.r && rowIndex <= range.e.r;
    }
};