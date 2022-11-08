function EnviadorDeMails() {
    Logger.log('Ejecutando enviador de mails');
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();

    var mail = data[0].indexOf('Dirección de correo electrónico');
    var puntaje = data[0].indexOf('Puntuación');
    var enviado = data[0].indexOf('');
    var enviado = data[0].indexOf('Nombre');
    var enviado = data[0].indexOf('Teléfono');
    var enviado = data[0].indexOf('Apellido');

    const niveles = [
        'Nivel Principiante 1',
        'Nivel Principiante 1',
        'Nivel Principiante 1',
        'Nivel Principiante 1',
        'Nivel Principiante 1',
        'Nivel Principiante 1',
        'Nivel Principiante 2',
        'Nivel Principiante 2',
        'Nivel Principiante 2',
        'Nivel Principiante 2',
        'Nivel Principiante 2',
        'Nivel Intermedio 1',
        'Nivel Intermedio 1',
        'Nivel Intermedio 1',
        'Nivel Intermedio 1',
        'Nivel Intermedio 1',
        'Nivel Intermedio 2',
        'Nivel Intermedio 2',
        'Nivel Intermedio 2',
        'Nivel Intermedio 2',
        'Nivel Intermedio 2',
        'Nivel Avanzado 1',
        'Nivel Avanzado 1',
        'Nivel Avanzado 1',
        'Nivel Avanzado 1',
        'Nivel Avanzado 1',
        'Nivel Avanzado 2',
        'Nivel Avanzado 2',
        'Nivel Avanzado 2',
        'Nivel Avanzado 2',
        'Nivel Avanzado 2',
    ]

    var globalRow;
    for (var i = 1; i < data.length; i++) {
        globalRow = data[i];
        if (get('Mail enviado') != "si") {
            emailReport();
            //sheet.getRange(1, i, data[0].indexOf('Mail enviado')).setValue("si")
            var a = sheet.getRange(i + 1, data[0].indexOf('Mail enviado') + 1, 1).setValue("si")
        }
    }

    function emailReport() {
        var fecha = new Date(get('Marca temporal'));
        var fecha_format = Utilities.formatDate(fecha, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
        var codigo = Utilities.formatDate(fecha, 'America/Argentina/Buenos_Aires', 'yyyyMMddHHmmssSS');

        var templ = HtmlService.createTemplateFromFile('email');
        templ.email = get('Dirección de correo electrónico');
        templ.nivel = niveles[get('Puntuación')];
        templ.telefono = get('Teléfono');
        templ.nombre = get('Nombre');
        templ.apellido = get('Apellido');
        templ.codigo = codigo;
        templ.fecha = fecha_format;
        var message = templ.evaluate().getContent();

        try {
            MailApp.sendEmail({
                to: 'idiomas@ihavemyamericanvisa.com',
                from: 'idiomas@ihavemyamericanvisa.com',
                replyTo: 'idiomas@ihavemyamericanvisa.com',
                cc: get('Dirección de correo electrónico'),
                subject: "Resultado test de ingles para visa",
                htmlBody: message
            });
        } catch (e) {
            // Logs an ERROR message.
            Logger.log(`Error: ${e}`);
        }

        Logger.log('Mail enviado');
    }

    function get(name) {
        var index = data[0].indexOf(name);
        return globalRow[index];
    }
}




