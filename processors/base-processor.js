class BaseProcessor {
    getName() {
        return this.constructor.name;
    }
    
    process(worksheet) {
        throw new Error('Process method must be implemented by subclass');
    }
    
    validateInput(worksheet) {
        return Validators.validateWorksheet(worksheet);
    }
    
    validateOutput(worksheet) {
        return Validators.validateWorksheet(worksheet);
    }
}