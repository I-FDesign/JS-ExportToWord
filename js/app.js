
let exportType = 'demanda';

let data = {};

$(document).ready(() => {
    loadForm(exportType);
    getDataFromStorage();

    const docoments = document.getElementsByClassName('to-export');
    hideAllElements(docoments);
})

function getDataFromStorage() {
    if(localStorage.getItem('data')) {
        data = JSON.parse(localStorage.getItem('data'));
    }
}

function hideAllElements(elements) {
    for(let i = 0; i < elements.length; i++) {
        $(elements[i]).css({
            'display': 'none'
        });
    }
}

function loadForm(formName) {
    const forms = document.getElementsByClassName('form');

    hideAllElements(forms);

    $('#' + formName + '_form').css({
        'display': 'block'
    });

    setTimeout(() => {
        setFormData();
    }, 200)
}

function loadDoc(docName) {
    const docs = document.getElementsByClassName('to-export');

    hideAllElements(docs);

    $('#' + docName + '_doc').css({
        'display': 'block'
    });

    const voutoTable = document.getElementById('vuoto_table');

    if(docName !== 'demanda') {
        voutoTable.style.display = 'none';
    } else {
        voutoTable.style.display = 'table';
    }

    setTimeout(() => {
        setFormData();
    }, 200)
}

function setFormData() {
    const formChildNodes = $('#' + exportType + '_form form')[0].getElementsByClassName('form-control');
    for(let i = 0; i < formChildNodes.length; i++) {
        const input = formChildNodes[i];
        input.value = data[input.name] ? data[input.name] : '';
    }
}

function resetData(){
    data = {};
    localStorage.removeItem('data');
    setFormData();
}

function changeType(type) {
    $('#' + exportType + '_btn').removeClass('active');
    exportType = type;
    loadForm(exportType);
    $('#' + type + '_btn').addClass('active');
}

function Export2Doc(data){
    var preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
    var postHtml = "</body></html>";
    loadDoc(exportType);
    setTimeout(() => {
        replaceInDoc().then(() => generateDocument(preHtml, postHtml, exportType));
    }, 1000)
}

function setDocDataAndExport() {
    const formData = $('#' + exportType + '_form form').serializeArray().reduce(function(obj, item) {
        obj[item.name] = item.value;
        return obj;
    }, {});

    Object.keys(formData).forEach(key => {
        if(!data[key] || data[key] !== formData[key]) {
            data[key] = formData[key];
        }
    });

    localStorage.setItem('data', JSON.stringify(data));

    Export2Doc(data);
}

function replaceInDoc() {
    return new Promise((resolve, reject) => {
        const doc = document.getElementById('element_to_export');
        const elements = doc.getElementsByClassName('field');

        for(let i = 0; i < elements.length; i++) {
            elements[i].innerHTML = data[elements[i].classList[0]];
        }

        setTimeout(() => {
            resolve();
        }, 500);
    })
}

function generateDocument(preHtml, postHtml, filename){
    var html = preHtml+document.getElementById('element_to_export').innerHTML+postHtml;

    var blob = new Blob(['\ufeff', html], {
        type: 'application/msword'
    });
    
    // Specify link url
    var url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(html);
    
    // Specify file name
    filename = filename?filename+'.doc':'document.doc';
    
    // Create download link element
    var downloadLink = document.createElement("a");

    document.body.appendChild(downloadLink);
    
    if(navigator.msSaveOrOpenBlob ){
        navigator.msSaveOrOpenBlob(blob, filename);
    }else{
        // Create a link to the file
        downloadLink.href = url;
        
        // Setting the file name
        downloadLink.download = filename;
        
        //triggering the function
        downloadLink.click();
    }

    document.body.removeChild(downloadLink);
}