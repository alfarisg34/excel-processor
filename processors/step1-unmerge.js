class Step1Unmerge extends BaseProcessor {
    getName() {
        return "Unmerge all cells";
    }
    
    process(worksheet) {
        if (!this.validateInput(worksheet)) {
            throw new Error('Invalid input for Step1Unmerge');
        }
        
        // Remove all merge ranges
        if (worksheet['!merges']) {
            worksheet['!merges'] = [];
        }
        
        if (!this.validateOutput(worksheet)) {
            throw new Error('Invalid output from Step1Unmerge');
        }
        
        return worksheet;
    }
}