
function Export2Doc(element, filename = '', data){
    var preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
    var postHtml = "</body></html>";
    $('#element_to_export').load('documents/1.html #1');
    setTimeout(() => {
        replaceInDoc(data).then(() => {
            generateDocument(preHtml, postHtml, filename);
        })
    }, 1000)
}

function setDocDataAndExport() {
    const data = $('#form_to_export').serializeArray().reduce(function(obj, item) {
        obj[item.name] = item.value;
        return obj;
    }, {});

    Export2Doc('1', 'testing', data);
}

function docWrite(varible) {
    document.write(variable);
}

function setInfo(className, value) {
    const elements = document.getElementsByClassName(className);

    for(let i = 0; i < elements.length; i++) {
        elements[i].innerHTML = value;
    }
}

function replaceInDoc(data) {
    return new Promise((resolve, reject) => {
        setInfo('nombre_cliente', data.nombre_cliente);
        setInfo('dni_cliente', data.dni_cliente);
        setInfo('nacionalidad_cliente', data.nacionalidad_cliente);
        setInfo('nacimiento_cliente', data.nacimiento_cliente);
        setInfo('profesion_cliente', data.profesion_cliente);
        setInfo('ingreso_anual_cliente', data.ingreso_anual_cliente);
        setInfo('domicilio_cliente', data.domicilio_cliente);
        setInfo('localidad_cliente', data.localidad_cliente);
        setInfo('partido_cliente', data.partido_cliente);
        setInfo('provincia_cliente', data.provincia_cliente);
        setInfo('marca_modelo_cliente', data.marca_modelo_cliente);
        setInfo('nombre_tercero', data.nombre_tercero);
        setInfo('dni_tercero', data.dni_tercero);
        setInfo('domicilio_tercero', data.domicilio_tercero);
        setInfo('localidad_tercero', data.localidad_tercero);
        setInfo('partido_tercero', data.partido_tercero);
        setInfo('marca_modelo_tercero', data.marca_modelo_tercero);
        setInfo('dominio_vehiculo_tercero', data.dominio_vehiculo_tercero);
        setInfo('aseguradora', data.aseguradora);
        setInfo('domicilio_aseguradora', data.domicilio_aseguradora);
        setInfo('fecha_accidente', data.fecha_accidente);
        setInfo('hora_accidente', data.hora_accidente);
        setInfo('lugar_accidente', data.lugar_accidente);
        setInfo('calle_circulaba_cliente', data.calle_circulaba_cliente);
        setInfo('direccion_circulaba_cliente', data.direccion_circulaba_cliente);
        setInfo('calle_circulaba_tercero', data.calle_circulaba_tercero);
        setInfo('direccion_circulaba_tercero', data.direccion_circulaba_tercero);
        setInfo('parte_cliente_afectada', data.parte_cliente_afectada);
        setInfo('parte_tercero_afectada', data.parte_tercero_afectada);
        setInfo('mecanica_accidente', data.mecanica_accidente);
        setInfo('intervencion_policial', data.intervencion_policial);
        setInfo('maniobra_esperable_tercero', data.maniobra_esperable_tercero);
        setInfo('lugar_atencion_cliente', data.lugar_atencion_cliente);
        setInfo('practicas_medicas_cliente', data.practicas_medicas_cliente);
        setInfo('secuelas_fisicas_cliente', data.secuelas_fisicas_cliente);
        setInfo('tratamiento_medico_cliente', data.tratamiento_medico_cliente);
        setInfo('porcentaje_incapacidad_fisica', data.porcentaje_incapacidad_fisica);
        setInfo('porcentaje_incapacidad_psiquica', data.porcentaje_incapacidad_psiquica);
        setInfo('monto_reclamado_incapacidad_fisica', data.monto_reclamado_incapacidad_fisica);
        setInfo('monto_reclamado_incapacidad_psiquica', data.monto_reclamado_incapacidad_psiquica);
        setInfo('monto_dano_moral', data.monto_dano_moral);
        setInfo('monto_presupuesto', data.monto_presupuesto);
        setInfo('monto_desvalorizacion', data.monto_desvalorizacion);
        setInfo('monto_demanda', data.monto_demanda);
        setInfo('nombre_testigo_1', data.nombre_testigo_1);
        setInfo('dni_testigo_1', data.dni_testigo_1);
        setInfo('domicilio_testigo_1', data.domicilio_testigo_1);
        setInfo('nombre_testigo_2', data.nombre_testigo_2);
        setInfo('dni_testigo_2', data.dni_testigo_2);
        setInfo('domicilio_testigo_2', data.domicilio_testigo_2);
        setInfo('nombre_testigo_3', data.nombre_testigo_3);
        setInfo('dni_testigo_3', data.dni_testigo_3);
        setInfo('domicilio_testigo_3', data.domicilio_testigo_3);
        setInfo('nombre_testigo_4', data.nombre_testigo_4);
        setInfo('dni_testigo_4', data.dni_testigo_4);
        setInfo('domicilio_testigo_4', data.domicilio_testigo_4);
        setInfo('nombre_testigo_5', data.nombre_testigo_5);
        setInfo('dni_testigo_5', data.dni_testigo_5);
        setInfo('domicilio_testigo_5', data.domicilio_testigo_5);

        setTimeout(() => {
            console.log('terminado');
            resolve();
        }, 1500)
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