builder.add('inputs','excel', class extends builder.InputClass {

    _init(){

        // Execute Parent Init
        super._init();

        // Set Additional Properties
        this._properties.autocomplete = 'off';
        this._properties.multiple = false;
    }

    _input(){

        // Create Input
        return $(document.createElement('input')).attr({
            'id': this._component.id + '-input',
            'class': 'form-control',
            'name': this._properties.name,
            'autocomplete': this._properties.autocomplete,
            'type': 'file',
            'accept': '.xlsx,.xls,.csv,.tsv,.ods',
        });
    }

    _extend(){
        if (this._properties.multiple) {
            this._component.input.attr('multiple', true);
            this._component.input.attr('name', this._properties.name + '[]');
        }
    }

    icon(fileName) {
        let extension = fileName.split('.').pop().toLowerCase();
        switch(extension) {
            case 'pdf': return 'file-earmark-pdf';
            case 'doc':
            case 'docx': return 'file-earmark-word';
            case 'xls':
            case 'xlsx': return 'file-earmark-excel';
            case 'ppt':
            case 'pptx': return 'file-earmark-ppt';
            case 'zip':
            case 'rar': return 'file-earmark-zip';
            case 'jpg':
            case 'jpeg':
            case 'png':
            case 'gif':
            case 'tiff':
            case 'bmp':
            case 'webp': return 'file-earmark-image';
            case 'mp3':
            case 'wav':
            case 'wma':
            case 'ogg':
            case 'm4a': return 'file-earmark-music';
            case 'mp4':
            case 'avi':
            case 'mkv':
            case 'wmv':
            case 'mov': return 'file-earmark-play';
            case 'css':
            case 'less':
            case 'scss':
            case 'sass':
            case 'js':
            case 'json':
            case 'xml':
            case 'html':
            case 'htm':
            case 'php':
            case 'asp':
            case 'aspx':
            case 'jsp':
            case 'cfm': return 'file-earmark-code';
            case 'txt':
            case 'log':
            case 'csv':
            case 'tsv': return 'file-earmark-text';
            case 'msg':
            case 'eml': return 'envelope-at';
            default: return 'file-earmark';
        }
    }

    read(file) {

        // Set Self
        const self = this;

        // Create File Reader
        return new Promise((resolve, reject) => {
            let reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const dataURL = e.target.result;
                    const base64String = dataURL.split(',')[1];
                    const binaryString = atob(base64String);
                    const workbook = XLSX.read(binaryString, { type: 'binary', cellDates: true, cellNF: true, cellText: false });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: null, dateNF: 'yyyy-mm-dd hh:mm:ss' });
                    resolve({
                        name: file.name,
                        size: file.size,
                        type: file.type,
                        extension: file.name.split('.').pop().toLowerCase(),
                        content: e.target.result,
                        json: jsonData,
                        icon: self.icon(file.name)
                    });
                } catch (error) {
                    reject('Error reading file: ' + error.message);
                    self.invalid(error.message);
                }
            };
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }

    val(value = null){

        // Set Self
        const self = this;

        // Create File Readers
        return new Promise((resolve, reject) => {
            // Initialize Values
            let values = [];

            // Retrieve the files
            let files = this._component.input[0].files;

            // Check if any file is selected
            if (files.length === 0) {
                resolve(values);
                return;
            }

            // Read all files
            let promises = [];
            for (let i = 0; i < files.length; i++) {
                promises.push(self.read(files[i]));
            }

            // Resolve all promises
            Promise.all(promises).then(fileData => {
                resolve(fileData);
            }).catch(error => {
                reject(error);
            });
        });
    }
});
