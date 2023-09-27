const fileUpload_Input = document.querySelector('#upload-file'),
    convertBtn = document.querySelector('#convert-btn'),
    resultArea = document.querySelector('#result-area');

convertBtn.addEventListener('click', () => {

    resultArea.value = '';
    resultArea.style.display = 'none';

    const files = fileUpload_Input.files;

    if (files.length === 0) {
        alert('Please choose an excel file..');
        return;
    }

    const filename = files[0].name;
    const extension = filename.substring(filename.lastIndexOf('.')).toLowerCase();

    if (extension === '.xls' || extension === '.xlsx') {
        convertExcelToJSON(files[0]);
    } else {
        alert('File should be of -  .xls | .xlsx');
    }

});


function convertExcelToJSON(file) {
    try {

        const reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = (e) => {
            const data = e.target.result;
            const workBook = XLSX.read(data, { type: 'binary' });
            let result = {};

            workBook.SheetNames.forEach((sheetName) => {
                const rowObjectArray = XLSX.utils.sheet_to_row_object_array(workBook.Sheets[sheetName]);
                if (rowObjectArray.length > 0) result = rowObjectArray;
            });

            // Showing the result
            resultArea.value = JSON.stringify(result, null, 4);
            resultArea.style.display = 'block';
            fileUpload_Input.value = '';
        }

    } catch (_) {
        console.log('Something went wrong!');
    }
}