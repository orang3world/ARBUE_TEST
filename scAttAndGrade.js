/* CODIGO GENERADO PARA REALIZAR INFORMES Y SEGUIMIENTOS A LOS ALUMNOS DEL
 PROGRAMA AWS RE/START EN POTRERO DIGITAL  */

//-----------------------------------------------------------------------------

const numCourse = 'TEST'
const courseName = "ARBUE_" + numCourse
//const numStud = 45

const columnas = 109;
const systemDate = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd")
const update = Utilities.formatDate(new Date(), "GMT-3", "dd MMM HH:mm")
// Name's Folders
const sourceFolderName = 'FILE_INCOMING_' + numCourse
const storageFolderName = 'STORAGE_' + numCourse

const attendanceBackupFolder = 'ATTENDANCE_BACKUP_' + numCourse
const gradeBackupFolder = 'GRADE_BACKUP_' + numCourse
const labBackupFolder = 'LAB_BACKUP_' + numCourse
const restBackupFolder = 'REST_BACKUP_' + numCourse

// Name's Spreadsheets
const attendanceSpName = 'attendance_' + numCourse
const gradeSpName = 'grade_' + numCourse
const labSpName = 'lab_' + numCourse
const studentSpName = 'studentData_' + numCourse
const att_gradeSpName = 'attendance_and_grade_' + numCourse
const backofficeSpName = 'backoffice_' + numCourse
const htmlSpName = 'html_' + numCourse

// Folders Access
function folderAccess(folderName) {
    return DriveApp.getFoldersByName(folderName).next()
}

// Spreadsheet Access
function spA(spName) {
    return SpreadsheetApp
        .openById(DriveApp.getFilesByName(spName).next().getId())
}

// Sheet Access
function ssA(spName, ssName) {
    return SpreadsheetApp
        .openById(DriveApp.getFilesByName(spName).next().getId())
        .getSheetByName(ssName)
}

// File Access
function fileIdByName(fileName) {
    return DriveApp.getFilesByName(fileName).next().getId()
}

// Arrays and Object
const spNames = [attendanceSpName, gradeSpName, labSpName]
var sheetNames = []
var csvList = {}    // Object for csv File info
var csvNames = []   // Array for csv File names
var csvIds = []     // Array for csv File ids
var instructorName = {}
var instructorLinkedin = {}
var instructorCel = {}
var instructorRole = {}
var eList = []

// Regex
//example of csv name = /participant-20230502222213.csv/
const attendanceRegex = /participant/
//example of csv name = /2023-05-04T1109_Calificaciones-AWS_ARBUE13.csv/      
const gradeRegex = /Calificaciones/
//example of csv name = /RESTCUR-000001 ES 23732-1767-a0R4N00000G9qOGUAZ_Labtime_1683209549930.csv/
const labRegex = /Labtime/

// Add a custom menu to the active spreadsheet, including a separator and a sub-menu.
//-----------------------------------------------------------------------------
function onOpen() {
    //-----------------------------------------------------------------------------
    const ui = SpreadsheetApp.getUi()

    ui
        .createMenu('ARBUE')
        .addItem('CARGAR datos', 'renewingData')
        .addSeparator()
        .addItem('GENERAR backoffice', 'backoffice')
        .addSeparator()
        .addSubMenu(ui.createMenu('EMAIL')
            .addItem('GENERAR informe', 'undoneKc')
            .addSeparator()
            .addItem('VISTA-PREVIA', 'previewEmail')
            .addSeparator()
            .addItem('ENVIAR', 'informeAcademico')
            .addSeparator())
        .addToUi();
}

Logger.log(update)

//-----------------------------------------------------------------------------
function renewingData() {
    //-----------------------------------------------------------------------------

    fileProcess()
    updateAttendanceAndGrade()
}

//  -----------------------------------------------------------------------------
function fileProcess() {
    //-----------------------------------------------------------------------------
    //  MESSAGE FOR USER
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Inicio del procesamiento de los archivos .csv'
        , 'PROCESANDO...');

    var csvFiles = folderAccess(sourceFolderName).getFiles()

    while (csvFiles.hasNext()) {
        let csvFile = csvFiles.next()
        let csvFileId = csvFile.getId()
        let csvFileName = csvFile.getName()

        csvList[csvFileName] = csvFileId    // Object for csv File {name: id}
    }

    csvNames = Object.keys(csvList)   // Array for csv File names
    csvIds = Object.values(csvList)   // Array for csv File ids

    for (let i = 0; i < csvNames.length; i++) {

        console.log("file name : " + csvNames[i])

        var csvFileAccess = DriveApp.getFileById(csvIds[i])           // Access csv file

        // Variables for csv Files of attendance

        if ((csvNames[i].match(attendanceRegex)) != null) {
            var spName = attendanceSpName
            var regex = /(.*)(\d{4})-(\d{2})-(\d{2}).*/  // pattern for inside date
            var newString = "$2$3$4"
            var encoding = 'UTF-16'
            var delimiter = '\t'
            var targetFolderName = attendanceBackupFolder
            var csvData = csvFileAccess.getBlob().getDataAsString(encoding).valueOf()
            var csv = Utilities.parseCsv(csvData, delimiter);

            for (let j = 0; j < 5; j++) {
                if (typeof (csv[j + 1][1]) != undefined) {
                    var ssName = csv[j + 1][1].replace(regex, newString)

                    if (ssName != csvNames[i].replace(/(^.*-)(\d{8})(.*)/, "$2")) {
                        var csvFileAccess = DriveApp.getFileById(csvIds[i]).setName(csvNames[i]
                            .replace(/(^.*-)(\d{8})(.*)/, "$1" + ssName + "$3"))           // Rename csv file

                        console.log("old title: " + csvNames[i] + " new title: " + csvNames[i]
                            .replace(/(^.*-)(\d{8})(.*)/, "$1" + ssName + "$3"))
                    }
                    break
                }
            }
            // Variables for a csv Files of grade

        } else if ((csvNames[i].match(gradeRegex)) != null) {
            var spName = gradeSpName
            var regex = /(.*)(\d{4})-(\d{2})-(\d{2}).*/
            var newString = "$2$3$4"
            var encoding = 'UTF-8'
            var delimiter = ','
            var targetFolderName = gradeBackupFolder
            var csvData = csvFileAccess.getBlob().getDataAsString(encoding).valueOf()
            var csv = Utilities.parseCsv(csvData, delimiter);
            var ssName = csvNames[i].replace(regex, newString)
            // Variables for a csv Files of lab

        } else if ((csvNames[i].match(labRegex)) != null) {
            var spName = labSpName
            var regex = /.*/
            var newString = systemDate
            var encoding = 'UTF-8'
            var delimiter = ','
            var targetFolderName = labBackupFolder
            var csvData = csvFileAccess.getBlob().getDataAsString(encoding).valueOf()
            var csv = Utilities.parseCsv(csvData, delimiter);
            var ssName = csvNames[i].replace(regex, newString)

            // Variables for a other Files

        } else {
            var regex = "//"
            var newString = ""
            var spName = ""
            var targetFolderName = restBackupFolder
            var ssName = csvNames[i].replace(regex, newString)

        }

        var targetFolderAccess = folderAccess(targetFolderName)  // Access target folder

        if (spName != "") {

            var spAccess = spA(spName)
            var ssAccess = ssA(spName, ssName)

            if (!ssAccess) {                 // Create a Sheet if she not exist
                spAccess.insertSheet(ssName);
            } else {
                ssAccess.clearContents()
            }

            // Load Data of csv File in to sheet
            console.log(ssAccess.getName())
            var success = ssAccess.getRange(1, 1, csv.length, csv[0].length).setValues(csv);

            // If a load data is successly, moves csv file to backup folder
            if (success && targetFolderAccess) {
                csvFileAccess.moveTo(targetFolderAccess)
            }

        } else {
            csvFileAccess.moveTo(targetFolderAccess)
        }
    }
    //  MESSAGE FOR USER
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Archivos procesados:\n' + csvNames.toString()
        , 'PROCESAMIENTO CONCLUIDO', 3)

}

// Ordenamiento de las hojas de cada planilla
/*---------------------------------------------------------------------------*/
function updateAttendanceAndGrade() {
    /*---------------------------------------------------------------------------*/
    var arrayStData = []
    var arrayKcData = []

    //var ssGradeLast = spGradeAccess.getSheets()[1]
    var ssGradeLast = spA(gradeSpName).getSheets()[1]
    var arrayGradeData = ssGradeLast.getDataRange().getValues()
    var arrayKcTempData2 = arrayGradeData.slice(2)
    arrayKcTempData2.push(arrayGradeData[1])

    // student data
    for (let i = 0; i < arrayGradeData.length; i++) {
        var familyName = arrayGradeData[i][0].replace(/(^.*), (.*$)/, "$1")
        var firstName = arrayGradeData[i][0].replace(/(^.*), (.*$)/, "$2")
        var studentId = arrayGradeData[i][1]
        var studentEmail = arrayGradeData[i][3]
        arrayStData.push([[familyName], [firstName], [studentId], [studentEmail]])
    }
    arrayStData.splice(0, 2, [['family Name'], ['first Name'], ['student Id'], ['student Email']])
    arrayStData.pop()

    for (let i = 0; i < arrayKcTempData2.length; i++) {
        var row = arrayKcTempData2[i].slice(6)
        var newRow = []
        for (let j = 0; j < row.length; j++) {
            try {
                if (row[j] != '') {
                    let item = Number(row[j])
                    newRow.push(item)
                } else {
                    newRow.push(row[j])
                }
            }
            catch (e) {
                newRow.push(row[j])
            }
        }
        arrayKcData.push(newRow)
    }

    var numStud = arrayStData.slice(1).length

    for (let i = 0; i < spNames.length; i++) {
        //var spFullId = fileIdByName(spNames[i])
        var spFullAccess = spA(spNames[i])
        var sss = spFullAccess.getSheets()
        if (!spFullAccess.setActiveSheet(spFullAccess.getSheetByName("Report"))) {
            var ssReport = sss[0].setName("Report")
        } else {
            var ssReport = spFullAccess.setActiveSheet(spFullAccess.getSheetByName("Report"));
        }
        sheetNames = []
        for (let j = 0; j < sss.length; j++) {
            sheetNames.push(sss[j].getName());
        }
        sheetNames.sort().reverse();
        /*---------------------------------------------------------------------------*/
        console.log('sheetNames sort.reverse : ' + sheetNames)
        /*---------------------------------------------------------------------------*/
        for (let k = 0; k < sheetNames.length; k++) {
            spFullAccess.setActiveSheet(spFullAccess.getSheetByName(sheetNames[k]));
            spFullAccess.moveActiveSheet(k + 1);
        }
        spFullAccess.setActiveSheet(spFullAccess.getSheetByName("Report"));
        spFullAccess.moveActiveSheet(1);

        // delete student's data
        ssReport.getRange(1, 1, ssReport.getLastRow(), arrayStData[0].length).clearContent()
        // default student's data (Report sheet)
        ssReport.getRange(1, 1, arrayStData.length, arrayStData[0].length).setValues(arrayStData)
        // delete header of columns
        ssReport.getRange(1, 5, 1, ssReport.getLastColumn()).clearContent()
        // complete header of columns with sheet names
        sheetNames.push([''])
        sheetNames.splice(sheetNames.indexOf('Report'), 1)
        ssReport.getRange(1, 5, 1, sheetNames.length).setValues([sheetNames])
        /*---------------------------------------------------------------------------*/
        console.log('sheetNames with \'Report\' sheet:' + sheetNames)
        /*---------------------------------------------------------------------------*/
        //else {ssReport.getRange(1, 5, 1, sheetNames.length).setValues([sheetNames])}
    }
    //  MESSAGE FOR USER
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Inicio de la carga de nuevos datos'
        , 'ACTUALIZANDO esta planilla ...');

    var ssReportAttAccess = ssA(attendanceSpName, 'Report')
    var attRepData = ssReportAttAccess
        .getRange(1, 1, ssReportAttAccess.getLastRow(), ssReportAttAccess.getLastColumn())
        .getValues()

    var ssData = []
    var ssDateName = []
    var course = {}
    // ---------------------------------------------------------------- Array with names of all ss 
    // spreadsheet for ATTENDANCE
    var ssAll = spA(attendanceSpName).getSheets()

    for (let h = 0; h < (ssAll.length); h++) {
        ssDateName.push(ssAll[h].getName())
        //ssDateName.splice(0)
    }
    ssDateName.splice(ssDateName.indexOf('Report'), 1)
    /*---------------------------------------------------------------------------*/
    console.log('ssDateName ' + ssDateName)
    /*---------------------------------------------------------------------------*/
    /*
      var reportHeaderRange = ssReportAttAccess.getRange(1, 5, 1, ssDateName.length)
      reportHeaderRange.setValues([ssDateName])
    
      console.log(reportHeaderRange)
    */
    for (let i = 0; i < ssDateName.length; i++) {

        var emails = {}
        var attSsName = ssDateName[i] //name of a sheet
        /*---------------------------------------------------------------------------*/
        console.log('attSsName :' + attSsName)
        /*---------------------------------------------------------------------------*/
        var ssAccess = ssA(attendanceSpName, attSsName) // Sheet access by your name
        var ssAccessReport = ssA(attendanceSpName, 'Report') // Sheet access by your name

        // ------------------------------------------------------ Array with data in everyone sheet
        console.log(spNames[i])
        console.log(ssAccess.getName())
        /*---------------------------------------------------------------------------*/
        var ssData = ssAccess.getRange(2, 8, ssAccess.getLastRow() - 1, 3).getValues()

        for (let j = 0; j < ssData.length; j++) {

            var email = ssData[j][0]
            var sT = new Date(ssData[j][1]) / 60000 // minutes
            var eT = new Date(ssData[j][2]) / 60000 // minutes

            if (!emails[email]) {
                Object.assign(emails, { [email]: { [sT]: eT } })
            } else {
                Object.assign(emails[email], { [sT]: eT })
            }

        }

        Object.assign(course, { [attSsName]: emails })
        var reportEmails = ssAccessReport
            .getRange(2, 4, ssAccessReport.getLastRow() - 1, 1)
            .getValues()

        var clearColumn = ssAccessReport
            .getRange(2, i + 1 + 4, ssAccessReport.getLastRow() - 1, 1)
            .clearContent()


        for (let k = 0; k < reportEmails.length; k++) {

            var eMail = reportEmails[k][0]

            if (course[attSsName][eMail]) {
                var sTList = Object.keys(course[attSsName][eMail]).sort()
                ssAccessReport.getRange(k + 2, i + 1 + 4).setValue(1)
            } else {
                ssAccessReport.getRange(k + 2, i + 1 + 4).setValue(0)
            }

        }
    }

    var ssAttWbx = ssA(att_gradeSpName, 'ASIST-WEBEX')

    // WEBEX : delete old data of students
    ssAttWbx.getRange(1, 1, ssAttWbx.getLastRow(), arrayStData[0].length).clearContent()
    // WEBEX : loading new data of students
    ssAttWbx.getRange(1, 1, arrayStData.length, arrayStData[0].length).setValues(arrayStData)
    // WEBEX : delete old data of attendance
    ssAttWbx.getRange(1, 8, ssAttWbx.getLastRow(), ssAttWbx.getLastColumn()).clearContent()
    // WEBEX : loading new data of attendance
    ssAttWbx.getRange(1, 8, attRepData.length, attRepData[0].length).setValues(attRepData)
    /*---------------------------------------------------------------------------*/
    console.log('attRepData ' + attRepData)
    /*---------------------------------------------------------------------------*/
    var ssKc = ssA(att_gradeSpName, 'KC')

    // row of possiblePoints
    //var pPoints = ssKc.getRange(ssKc.getLastRow(), 1, 1, ssKc.getLastColumn()).getValues()

    //console.log('copy pPoints ' + pPoints)

    // KC : delete old data of students
    ssKc.getRange(1, 1, numStud, arrayStData[0].length).clearContent()
    // KC : loading new data of students
    ssKc.getRange(1, 1, arrayStData.length, arrayStData[0].length).setValues(arrayStData)
    // KC : delete old data of student's grades
    ssKc.getRange(2, 10, ssKc.getLastRow(), arrayKcData[0].length).clearContent()
    // KC : loading new data of student's grades
    ssKc.getRange(2, 10, arrayKcData.length, arrayKcData[0].length).setValues(arrayKcData)

    // insert copy possiblePoints row
    //ssKc.getRange(ssKc.getLastRow(), 1, 1, ssKc.getLastColumn()).setValues(pPoints)
    /*---------------------------------------------------------------------------*/
    //console.log('insert pPoints ' + pPoints)
    /*---------------------------------------------------------------------------*/
    var ssDkc = ssA(att_gradeSpName, 'd-kc')
    // d-kc : delete old data
    ssDkc.clearContents()
    // d-kc : loading new data
    ssDkc.getRange(1, 1, arrayGradeData.length, arrayGradeData[0].length).setValues(arrayGradeData)

    var ssGradeLastName = ssGradeLast.getName()
    ssKc.getRange(1, 2).setValue(ssGradeLastName)

    //  MESSAGE FOR USER
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Planilla lista para compartir. ( Nota : recuerde agregar los encabezados'
        + ' de los KC mas recientes)'
        , 'ACTUALIZACION TERMINADA', 7)
}
//-----------------------------------------------------------------------------
function backoffice() {
    //-----------------------------------------------------------------------------

    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Envio de copia de la informacion actualizada'
        , 'CONPARTIENDO backoffice');

    var sTarget = spA(backofficeSpName)
    var sSource = spA(att_gradeSpName)

    var ssSource = sSource.getSheets()
    var ssSource0 = ssSource[0]
    var rangeSource0 = ssSource0.getDataRange()
    var valuesSource0 = rangeSource0.getValues()

    ssSource0.copyTo(sTarget)

    var ssTarget = sTarget.getSheets()
    var ssTarget0 = ssTarget[0]
    var ssTarget1 = ssTarget[1]
    var ssTarget2 = ssTarget[2]
    var kcAsist = sTarget.getSheetByName("kc_asist_" + numCourse)
    var rangeTarget2 = ssTarget2.getDataRange()
    var newTarget2 = rangeTarget2.setValues(valuesSource0)
    sTarget.deleteSheet(kcAsist)
    ssTarget2.setName("kc_asist_" + numCourse)
    ssTarget0.getRange(1, 8).setValue("Update " + '\n' + update)
    ssTarget0.getRange(1, 1).setValue(courseName)

    //  MESSAGE FOR USER
    return (SpreadsheetApp.getActiveSpreadsheet().toast(
        'Se completo el envio de la informacion a la planilla Backoffice'
        , 'BACKOFFICE enviado')
    )
}

//=================================================================================
function undoneKc() {
    //-----------------------------------------------------------------------------
    //  MESSAGE FOR USER
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Recopilando informacion de asistencia y de KCs '
        , 'GENERANDO Informe Academico ...');

    var spGradeAccess = spA(gradeSpName)
    var ssKc = ssA(att_gradeSpName, 'KC')
    var numStud = spGradeAccess.getSheets()[1].getLastRow() - 3
    var ssUpdateKc = ssA(att_gradeSpName, 'UPDATE-KC')
    //-----------------------------------------------------------------------------
    console.log("numStud = " + numStud)
    //-----------------------------------------------------------------------------
    const headerRange = ssKc.getRange(1, 1, 1, columnas).getValues();
    const dataRange = ssKc.getRange(2, 1, numStud, columnas).getValues();
    const possiblePointsRange = ssKc.getRange(numStud + 3, 1, 1, columnas).getValues();

    var borrarDatos = ssUpdateKc
        .getRange(2, 1, ssUpdateKc.getLastRow(), ssUpdateKc.getLastColumn())
        .clearContent();

    dataRange.forEach(function (indice) {
        var cellEmail = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 1);
        var cellStudent = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 2);
        var cellAttendance = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 3);
        var cellKcOk = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 4);
        var cellState = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 5);
        var cellTOTALFALT = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 6);
        var cellNotDone = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 7);
        var cellNotDoneList = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 8);
        var cellLowGrades = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 9);
        var cellLowGradesList = ssUpdateKc.getRange(ssUpdateKc.getLastRow() + 1, 10);
        //-----------------------------------------------------------------------------
        var state = "";
        var numNotDoneKc = 0
        var numLowGradesKc = 0
        var col1KC = 10
        var pApproval = 0.7
        //-----------------------------------------------------------------------------
        var notDoneList = []
        var lowGradesList = []
        //-----------------------------------------------------------------------------
        cellEmail.setValue(indice[3]);
        cellStudent.setValue(indice[0] + " " + indice[1]);
        cellAttendance.setValue((indice[4] * 100).toFixed(0));
        cellKcOk.setValue((indice[5] * 100).toFixed(0));
        //-----------------------------------------------------------------------------
        //notDoneList.push('<ol>' + '\n')
        notDoneList.push('<ol>')
        for (i = col1KC; i <= columnas; i++) {
            headerRange.forEach(function (col) {
                if (indice[i - 1] == "" && col[i - 1] != "") {
                    numNotDoneKc++
                    console.log(numNotDoneKc)
                    //notDoneList.push('<li>' + col[i - 1] + '</li>' + '\n');
                    notDoneList.push('<li>' + col[i - 1] + '</li>');
                }
            })
        }
        notDoneList.push('</ol>')
        cellNotDoneList.setValue(notDoneList.join(''))
        //-----------------------------------------------------------------------------
        //lowGradesList.push('<ol>' + '\n')
        lowGradesList.push('<ol>')
        for (i = col1KC; i < columnas; i++) {
            headerRange.forEach(function (col) {
                if (indice[i - 1] != "" && col[i - 1] != "") {
                    possiblePointsRange.forEach(function (pointsP) {
                        if (indice[i - 1] < pointsP[i - 1] * pApproval) {
                            numLowGradesKc++
                            console.log(numNotDoneKc)
                            //lowGradesList.push('<li>' + col[i - 1] + '</li>' + '\n');
                            lowGradesList.push('<li>' + col[i - 1] + '</li>');
                        }
                    })
                }
            })
        }
        lowGradesList.push('</ol>')
        cellLowGradesList.setValue(lowGradesList.join(''))
        //-----------------------------------------------------------------------------
        if (numNotDoneKc + numLowGradesKc == 0) {
            var state = "AL DIA"; Logger.log(state);
        } else {
            var state = "FALTANTES"; Logger.log(state);
        }
        //-----------------------------------------------------------------------------
        cellState.setValue(state);
        cellNotDone.setValue(numNotDoneKc);
        cellLowGrades.setValue(numLowGradesKc);
        cellTOTALFALT.setValue(numLowGradesKc + numNotDoneKc);
    })
    //-----------------------------------------------------------------------------
    var cellUpdateDate = ssUpdateKc.getRange(2, 11);
    cellUpdateDate.setValue(new Date());
    //-----------------------------------------------------------------------------
    //  MESSAGE FOR USER
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Se han generado los informes individuales para cada estudiante'
        , 'INFORME Academico completado')
}

//=================================================================================
function emailAll() {
    //-----------------------------------------------------------------------------
    //  MESSAGE FOR USER
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Pasaje del INFOMRE a formato html');

    var eList = []
    var ssUpdateKc = ssA(att_gradeSpName, 'UPDATE-KC')
    /*
        const bodyData = ssBodyHtml.getDataRange().getValues()
        const signatureData = ssSignatureHtml.getDataRange().getValues()
        const instructorData = ssInstructorDataHtml.getDataRange().getValues()*/

    const bodyData = ssA(htmlSpName, 'BODY').getDataRange().getValues()
    const signatureData = ssA(htmlSpName, 'SIGNATURE').getDataRange().getValues()
    const instructorData = ssA(htmlSpName, 'INSTRUCTOR_DATA').getDataRange().getValues()

    var email_prof = Session.getActiveUser().getEmail();
    var instructorWorkData = instructorData.slice(1)  // with headers

    for (let i = 0; i < instructorWorkData.length; i++) {
        var iEmail = instructorWorkData[i][0]
        var iName = instructorWorkData[i][1]
        var iLink = instructorWorkData[i][2]
        var iCel = instructorWorkData[i][3]
        var iRole = instructorWorkData[i][4]

        Object.assign(instructorName, { [iEmail]: iName })
        Object.assign(instructorLinkedin, { [iEmail]: iLink })
        Object.assign(instructorCel, { [iEmail]: iCel })
        Object.assign(instructorRole, { [iEmail]: iRole })
    }

    var signature = signatureData[0][1]
    signature = signature.replace("{{instructorName[email_prof]}}", instructorName[email_prof])
    signature = signature.replace("{{instructorLinkedin[email_prof]}}", instructorLinkedin[email_prof])
    signature = signature.replace("{{instructorCel[email_prof]}}", instructorCel[email_prof])
    signature = signature.replace("{{instructorRole[email_prof]}}", instructorRole[email_prof])

    /* array datas in update-kc ----*/
    var updateKcData = ssUpdateKc.getRange(2, 1, ssUpdateKc.getLastRow() - 1, ssUpdateKc.getLastColumn()).getValues()
    var variableName = ["{{email}}", "{{student}}", "{{attendance}}", "{{kcok}}", "{{state2}}", "{{totalfalt}}", "{{notDone}}", "{{notDoneKcList}}", "{{lowGrade}}", "{{lowGradeKcList}}"]

    /* iterate through array rows --*/
    updateKcData.forEach(

        function crearMensaje(value) {
            var body = bodyData[0][1]
            var variable = {}
            /*VARIABLES DENTRO DEL E-MAIL */
            for (let i = 0; i < variableName.length; i++) {
                if (i == 1) {
                    Object.assign(variable, { "{{student}}": value[i].toString().toUpperCase() })
                    body = body.replace([variableName[i]], value[i].toString().toUpperCase())
                } else {
                    Object.assign(variable, { [variableName[i]]: value[i] })
                    body = body.replace([variableName[i]], value[i])
                }
            }
            body = body.replace("{{signature}}", signature)

            /* utilizar https://codepen.io/ para el armado del HTML */
            /* usar https://www.textfixer.com/html/compress-html-compression.php para generar html en una sola linea */

            var new_subject = "Informe Académico " + variable["{{student}}"] + " - AWS re/Start " + courseName;
            var empty_msj = "";
            var ccemail = "ariel.orange@compromiso.org"

            eList.push({
                'to': variable["{{email}}"],
                'cc': ccemail,
                'subject': new_subject,
                'body': empty_msj,
                'htmlBody': body
            })

        }
    )
    return eList
}
//=================================================================================
function previewEmail() {
    //-----------------------------------------------------------------------------
    /*  MESSAGE FOR USER
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Inicio del envio de e-mails personalizados '
        , 'ENVIANDO Informe Academico ...');
*/
    var ui = SpreadsheetApp.getUi()

    var eList = emailAll()
    var result = ui
        .prompt('Indique la vista previa que desea ver por medio de la fila', 'VISTA PREVIA', ui.ButtonSet.OK)
    var text = result.getResponseText()
    var html = HtmlService.createHtmlOutput(eList[text - 2].htmlBody)
    //ui.showsi(html, 'Email - Preview')
    ui.showSidebar(html)

    /*  MESSAGE FOR USER
    return (SpreadsheetApp.getActiveSpreadsheet().toast(
        'todos los informes academicos han sido enviados '
        , 'FINALIZADO el envio de e-mails'))
        */
}
//=================================================================================
function informeAcademico() {
    //-----------------------------------------------------------------------------
    undoneKc()
    previewEmail()
    var eList = emailAll()

    var ui = SpreadsheetApp.getUi()
    var result = ui.alert(
        'Informe Academico',
        'Esta a punto de enviar '
        + eList.length
        + ' emails.\n Esta acción no puede deshacerse.\n ¿ Desea enviar los emails ?',
        ui.ButtonSet.YES_NO);
    var button = result.getSelectedButton();
    //What happens if the User clicks OK
    if (button == ui.Button.YES) {
        for (let i = 0; i < eList.length; i++) {
            /*
                        MailApp.sendEmail(elist[i]);
                        */
        }
    }

}