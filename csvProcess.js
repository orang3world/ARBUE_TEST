/* CODIGO GENERADO PARA REALIZAR INFORMES Y SEGUIMIENTOS A LOS ALUMNOS DEL
 PROGRAMA AWS RE/START EN POTRERO DIGITAL  */

//-----------------------------------------------------------------------------

const numCourse = 'TEST'
const courseName = "ARBUE_" + numCourse
//const numStud = 45

const columnas = 109;
const systemDate = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd")
const update = Utilities.formatDate(new Date(), "GMT-3", "dd MMM HH:mm")

// Names of incoming Folders
const sourceFolderName = 'FILE_INCOMING_' + numCourse
const storageFolderName = 'STORAGE_' + numCourse

// Names of target Folders
const attendanceBackupFolder = 'ATTENDANCE_BACKUP_' + numCourse
const gradeBackupFolder = 'GRADE_BACKUP_' + numCourse
const labBackupFolder = 'LAB_BACKUP_' + numCourse
const restBackupFolder = 'REST_BACKUP_' + numCourse

// Names of Spreadsheets
const attendanceSpName = 'attendance_' + numCourse
const gradeSpName = 'grade_' + numCourse
const labSpName = 'lab_' + numCourse
const studentSpName = 'studentData_' + numCourse
const att_gradeSpName = 'attendance_and_grade_' + numCourse
const backofficeSpName = 'backoffice_' + numCourse
const htmlSpName = 'html_' + numCourse

// Access 
const csvFolderAccess = DriveApp.getFoldersByName(sourceFolderName).next()
const csvFiles = csvFolderAccess.getFiles()

const studentDataFileAccess = DriveApp.getFilesByName(studentSpName).next()
const studentDataFileId = studentDataFileAccess.getId()
const spStDataAccess = SpreadsheetApp.openById(studentDataFileId)
const numStud = spStDataAccess.getLastRow() - 1
const arrayStData = spStDataAccess.getSheets()[0].getRange(1, 1, spStDataAccess.getLastRow
  (), 4).getValues()

const spGradeFileAccess = DriveApp.getFilesByName(gradeSpName).next()
const spGradeFileId = spGradeFileAccess.getId()
const spGradeAccess = SpreadsheetApp.openById(spGradeFileId)

const spHtmlFileAccess = DriveApp.getFilesByName(htmlSpName).next()
const spHtmlFileId = spHtmlFileAccess.getId()
const spHtmlAccess = SpreadsheetApp.openById(spHtmlFileId)
const ssBodyHtml = spHtmlAccess.getSheetByName('BODY')
const ssSignatureHtml = spHtmlAccess.getSheetByName('SIGNATURE')
const ssInstructorDataHtml = spHtmlAccess.getSheetByName('INSTRUCTOR_DATA')
const dataRangeBody = ssBodyHtml.getDataRange().getValues()
const dataRangeSignature = ssSignatureHtml.getDataRange().getValues()
const dataRangeInstructorData = ssInstructorDataHtml.getDataRange().getValues()

const spBackofficeFileAccess = DriveApp.getFilesByName(backofficeSpName).next()
const spBackofficeFileId = spBackofficeFileAccess.getId()
const spBackofficeAccess = SpreadsheetApp.openById(spBackofficeFileId)

const spAttGradeAccess = SpreadsheetApp.openById(DriveApp.getFilesByName(att_gradeSpName).next().getId())
const ssDkc = spAttGradeAccess.getSheetByName('d-kc')
const ssKc = spAttGradeAccess.getSheetByName('KC')
const ssUpdateKc = spAttGradeAccess.getSheetByName('UPDATE-KC');

const spAttAccess = SpreadsheetApp.openById(DriveApp.getFilesByName(attendanceSpName).next().getId())
const ssReportAccess = spAttAccess.getSheetByName("Report")
const rangeDataRepor = ssReportAccess.getRange(1, 5, ssReportAccess.getLastRow(), ssReportAccess.getLastColumn() - 4).getValues()
const ssAll = spAttAccess.getSheets()

// Ranges
const headerRange = ssKc.getRange(1, 1, 1, columnas).getValues();
const dataRange = ssKc.getRange(2, 1, numStud, columnas).getValues();
const possiblePointsRange = ssKc.getRange(numStud + 3, 1, 1, columnas).getValues();

// Arrays and Object
const spNames = [attendanceSpName, gradeSpName, labSpName]
var sheetNames = []
var csvList = {}    // Object for csv File info
var csvNames = []   // Array for csv File names
var csvIds = []     // Array for csv File ids
var instructorName = {}
var instructorLinkedin = {}

// Regex
//example of csv name = /participant-20230502222213.csv/
const attendanceRegex = /participant/
//example of csv name = /2023-05-04T1109_Calificaciones-AWS_ARBUE13.csv/      
const gradeRegex = /Calificaciones/
//example of csv name = /RESTCUR-000001 ES 23732-1767-a0R4N00000G9qOGUAZ_Labtime_1683209549930.csv/
const labRegex = /Labtime/
//-----------------------------------------------------------------------------

// Add a custom menu to the active spreadsheet, including a separator and a sub-menu.
//-----------------------------------------------------------------------------
function onOpen() {
  //-----------------------------------------------------------------------------

  SpreadsheetApp.getUi()
    .createMenu('ARBUE')
    .addItem('CARGAR datos', 'renewingData')
    .addSeparator()
    .addItem('GENERAR backoffice', 'backoffice')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('EMAIL')
      .addItem('GENERAR informe', 'undoneKc')
      .addSeparator()
      .addItem('ENVIAR', 'informeAcademico')
      .addSeparator())
    .addToUi();
}


Logger.log(update)

//-----------------------------------------------------------------------------
function backoffice() {
  //-----------------------------------------------------------------------------

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Envio de copia de la informacion actualizada'
    , 'CONPARTIENDO backoffice');

  var sTarget = spBackofficeAccess
  var sSource = spAttGradeAccess

  var ssSource = sSource.getSheets()
  var ssSource0 = ssSource[0]
  var rangeSource0 = ssSource0.getDataRange()
  var valuesSource0 = rangeSource0.getValues()

  ssSource0.copyTo(sTarget)

  var ssTarget = sTarget.getSheets()
  var ssTarget0 = ssTarget[0]
  var ssTarget1 = ssTarget[1]
  var ssTarget2 = ssTarget[2]
  var kcAsist = sTarget.getSheetByName("kc_asist")
  var rangeTarget2 = ssTarget2.getDataRange()
  var newTarget2 = rangeTarget2.setValues(valuesSource0)
  sTarget.deleteSheet(kcAsist)
  ssTarget2.setName("kc_asist")
  ssTarget0.getRange(1, 8).setValue("Update " + '\n' + update)
  ssTarget0.getRange(1, 1).setValue(courseName)

  return (SpreadsheetApp.getActiveSpreadsheet().toast(
    'Se completo el envio de la informacion a la planilla Backoffice'
    , 'BACKOFFICE enviado')
    )
}

//-----------------------------------------------------------------------------
function undoneKc() {
  //-----------------------------------------------------------------------------

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Recopilando informacion de asistencia y de KCs '
    , 'GENERANDO Informe Academico ...');

  var borrarDatos = ssUpdateKc.getRange(2, 1, ssUpdateKc.getLastRow(), ssUpdateKc.getLastColumn()).clearContent();
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
    cellEmail.setValue(indice[3]);
    cellStudent.setValue(indice[0]+" "+indice[1]);
    cellAttendance.setValue((indice[4] * 100).toFixed(0));
    cellKcOk.setValue((indice[5] * 100).toFixed(0));
    //-----------------------------------------------------------------------------
    cellNotDoneList.setValue(cellNotDoneList.getValue() + '<ol>' + '\n')
    for (i = col1KC; i <= columnas; i++) {
      headerRange.forEach(function (col) {
        if (indice[i - 1] == "" && col[i - 1] != "") {
          numNotDoneKc++
          console.log(numNotDoneKc)
          cellNotDoneList.setValue(cellNotDoneList.getValue() + '<li>' + col[i - 1] + '</li>' + '\n');
        }
      })
    }

    cellNotDoneList.setValue(cellNotDoneList.getValue() + '</ol>')

    //-----------------------------------------------------------------------------
    cellLowGradesList.setValue(cellLowGradesList.getValue() + '<ol>' + '\n')
    for (i = col1KC; i < columnas; i++) {
      headerRange.forEach(function (col) {
        if (indice[i - 1] != "" && col[i - 1] != "") {
          possiblePointsRange.forEach(function (pointsP) {
            if (indice[i - 1] < pointsP[i - 1] * pApproval) {
              numLowGradesKc++
              console.log(numNotDoneKc)
              cellLowGradesList.setValue(cellLowGradesList.getValue() + '<li>' + col[i - 1] + '</li>' + '\n');
            }
          })
        }
      })
    }
    cellLowGradesList.setValue(cellLowGradesList.getValue() + '</ol>')

    if (numNotDoneKc + numLowGradesKc == 0) {
      var state = "AL DIA"; Logger.log(state);
    } else {
      var state = "FALTANTES"; Logger.log(state);
    }

    cellState.setValue(state);
    cellNotDone.setValue(numNotDoneKc);
    cellLowGrades.setValue(numLowGradesKc);
    cellTOTALFALT.setValue(numLowGradesKc + numNotDoneKc);
  })

  var cellUpdateDate = ssUpdateKc.getRange(2, 11);
  cellUpdateDate.setValue(new Date());

  return (SpreadsheetApp.getActiveSpreadsheet().toast(
    'Se han generado los informes individuales para cada estudiante'
    , 'INFORME Academico completado'))
}

//-----------------------------------------------------------------------------
function informeAcademico() {
  //-----------------------------------------------------------------------------
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Inicio del envio de e-mails personalizados '
    , 'ENVIANDO Informe Academico ...');
  /* array datas in update-kc ----*/
  var dataRangeUPDATEKC = ssUpdateKc.getRange(2, 1, ssUpdateKc.getLastRow() - 1, ssUpdateKc.getLastColumn()).getValues()

  /* iterate through array rows --*/
  dataRangeUPDATEKC.forEach(

    function crearMensaje(value) {
      /* VARIABLES DENTRO DEL E-MAIL */

      var email = value[0];
      var ccemail = "aaorange75@gmail.com"// "potrerodigital@compromiso.org";
      var student = value[1];
      var attendance = value[2];
      var kcok = value[3];
      var state2 = value[4];
      var totalfalt = value[5];
      var notDone = value[6];
      var notDoneKcList = value[7];
      var lowGrade = value[8];
      var lowGradeKcList = value[9];

      var arrayInstructorData = dataRangeInstructorData.slice(1)
      var email_prof = Session.getActiveUser().getEmail();

      for (let i = 0; i < arrayInstructorData.length; i++) {
        var iEmail = arrayInstructorData[i][0]
        var iName = arrayInstructorData[i][1]
        var iLink = arrayInstructorData[i][2]

        Object.assign(instructorName, { [iEmail]: iName })
        Object.assign(instructorLinkedin, { [iEmail]: iLink })

      }

      //var signature = dataRangeSignature[0][0]
      //var body_html = HtmlService.createHtmlOutput(dataRangeBody[0][0])

      /* utilizar https://codepen.io/ para el armado del HTML */

      var signature = '<table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr class="signature" style=""> <td valign="top" align="center" class="center-on-narrow stack-column" style="width: 100%; overflow: hidden;"> <div style="border: 0px solid transparent;color: #f1f4f6; display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.2; Margin: 25px 20px 25px 25px; max-height: none; max-width: none; padding: 0px; text-decoration: none;"> <div> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="4" class="nombre"> <b>' + instructorName[email_prof] + '</b> | Instructor </font> </div> <div> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="3">AWS Certified Associate<br></font> <div class=""> <font size="3"><b> <font face="&quot;Arial&quot;, Helvetica, sans-serif"> <a href="http://compromiso.org/" style="color: #ec7211;" class="">compromiso.org</a> </font> </b></font> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="3">| </font> <font size="3"><b> <font face="&quot;Arial&quot;, Helvetica, sans-serif"> <a href="http://potrerodigital.org/" class="" style="color: #ec7211;">potrerodigital.org</a> </font> </b></font> <br> </div> <div class="firm"> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="3"> Cel. +54 9 11 5895-3808</font> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="3">| </font> <font size="2"><b> <font face="&quot;Arial&quot;, Helvetica, sans-serif"> <a href="' + instructorLinkedin[email_prof] + '" class="" style="color: #ec7211;">LinkedIn</a> </font> </b></font> <br> </div> </div> </div> </td> </tr> <tr class="signature"> <td valign="top" align="center" class="center-on-narrow stack-column" style="width: 100%; overflow: hidden; vertical-align: bottom; padding-bottom: 15px;"> <img src="https://ci4.googleusercontent.com/proxy/ZDZm9x8NP-OYdpemo9sXn8iq8qc7K4FslWZBBTo1zg21pCjv13Ph6KMzbIU0g2oAkEU71HA-gMsWoqGwPYj7vWpnD7xZARpNotH0BnphC7aMccAc7618dQO8o_dLzeDBDGu2yiynJvrX1or5KzfM=s0-d-e1-ft#https://images.credly.com/size/340x340/images/00634f82-b07f-4bbd-a6bb-53de397fc3a6/image.png" alt="AWS Certified Cloud Practitioner" width="96" height="96" class="CToWUd" data-bit="iit"> <img src="https://ci3.googleusercontent.com/proxy/VpBzkJjDHaXcQzbidXDpRdPSPahgJpbU0hDNcIlrLBYVjLsxUevGLzYdvNk6KUL8LCGpejdGVl6U02zuSjo3ga91sTAokXYo1Tm1Z3t0Kc1p6h4xg6shcgYuMVmy_rP_SmFkXtreax1p1qLovE3N=s0-d-e1-ft#https://images.credly.com/size/340x340/images/0e284c3f-5164-4b21-8660-0d84737941bc/image.png" alt="AWS Certified Solutions Architect – Associate" width="96" height="96" class="CToWUd" data-bit="iit"> <img src="https://ci3.googleusercontent.com/proxy/ep9w9gvBrTwl4kJ19bQf0B3BaFV-O9Bfd1ooVx2pJWVong0E4Sxa2NdRdZ5Atfs36cC13_SfW3IeTqCmO9lneyty4VFwiSvX7QC096MWg_sDU0o8-EvZBX1FWgbS5FH3q1DWlDqudPGTeUEpNpiM=s0-d-e1-ft#https://images.credly.com/size/340x340/images/44e2c252-5d19-4574-9646-005f7225bf53/image.png" alt="AWS re/Start Graduate" width="96" height="96" class="CToWUd" data-bit="iit"> <img src="https://ci6.googleusercontent.com/proxy/R-aIAR8Uu3ffYUJisbl6bfL1YWUP8r_JXip-EIbcQV-er5fz-wbsvK4HA7qItSyDrVV70pze-9hgT6MPwt2pfkYAIUtPwNtIASl_1MxA1MCaVn4soT1EGZpv5XsakDHw8PqTsVXIvbFtVLfQaZqw=s0-d-e1-ft#https://images.credly.com/size/340x340/images/e426d40e-8a6a-4f72-866e-2abfcfbde46b/image.png" alt="AWS re/Start Accredited Instructor" width="96" height="96" class="CToWUd" data-bit="iit"> </td> </tr> </tbody> </table>';
      var body_html = HtmlService.createHtmlOutput('<html><head> <meta charset="utf-8"> <meta name="viewport" content="width=device-width"> <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"> <meta http-equiv="X-UA-Compatible" content="IE=edge"> <meta name="robots" content="noindex"> <base target="_blank"> <style type="text/css"> body, div[style*="margin: 16px 0"], html { margin: 0 !important } body, html { padding: 0 !important; height: 100% !important; width: 100% !important } * { -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100% } table, td { mso-table-lspace: 0 !important; mso-table-rspace: 0 !important } table { border-spacing: 0 !important; border-collapse: collapse !important; margin: 0 auto !important } table table table { table-layout: auto } img { -ms-interpolation-mode: bicubic } .yshortcuts to { border-bottom: none !important } .mobile-link--footer a, a[x-apple-data-detectors] { color: inherit !important; text-decoration: underline !important } .signature{ background-color: #232f3e; } @media screen and (max-width:600px) { .stack-column-half { width: 50% !important; display: inline-block !important } .center-on-narrow, .fluid, .fluid-centered { margin-left: auto !important; margin-right: auto !important } table { table-layout: fixed !important } .email-container { width: 100% !important } .fluid, .fluid-centered { max-width: 100% !important; height: auto !important } .stack-column, .stack-column-center, .stack-column-full-width { display: block !important; width: 100% !important; max-width: 100% !important; direction: ltr !important } .center-on-narrow { display: block !important; float: none !important } table.center-on-narrow { display: inline-block !important } .stack-column-full-width .eddie-wrapper { color: white; width: 100% } } </style></head><body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" offset="0"> <table cellspacing="0" cellpadding="0" border="0" width="100%" style="font-family: Helvetica, Arial, sans-serif; width: 100%; padding: 20px; background-color: rgb(235, 235, 235); background-image: none;"> <tbody> <tr> <td align="center"> <table class="email-container" width="660"> <tbody> <tr> <td> <div class="eddie-page"> <!-- barra inicial --> <table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr style="background-color: #232f3e;"> <td valign="top" align="center" stackclass="stack-column-full-width" class="stack-column-full-width" style="width: 50%; overflow: hidden;"> <img src="https://d1.awsstatic.com/training-and-certification/Logos/aws_restart_logo_reverse.860113148166c4742ebd63e8fa74d09ae4cf64ea.png" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width:50%; color: rgb(0, 0, 0); display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 5px; max-height: none; max-width: none; padding: 5px; text-decoration: none; min-height: 10px;" alt=""> </td> <td valign="top" align="left" stackclass="stack-column-full-width" class="stack-column-full-width" style="width: 50%; overflow: hidden;"> <img src="https://static.wixstatic.com/media/5b90eb_2f1f983af79a4e69ba942bc0586dbb7d~mv2.png/v1/fill/w_382,h_40,al_c,q_85,usm_0.66_1.00_0.01,enc_auto/potrero_digital_2021_edited.png" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: 80%; color: rgb(0, 0, 0); display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 10px; max-height: none; max-width: none; padding: 10px; text-decoration: none; min-height: 10px;" alt=""> </td> </tr> </tbody> </table> <!-- imagen grande --> <table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr style="background-color: #ec7211;"> <td valign="top" class="center-on-narrow stack-column" style="width: 100%; overflow: hidden;"> <div style="border: 0px solid transparent;color: #ec7211; display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.2; Margin: 25px 20px 25px 25px; max-height: none; max-width: none; padding: 0px; text-decoration: none;"> <div style="text-align:center;"> <font color="white" face="Arial, Helvetica, sans-serif" size="7"> <b>Informe Académico</b> </font> </div> </div> </td> </tr> </tbody> </table> <!-- cuerpo --> <table cellspacing="0" cellpadding="0" border="0" style="width: 700px; height: 528px;"> <tbody> <tr style="background-color: #ced2d5;"> <td valign="top" class="center-on-narrow stack-column" style="width: 100%; overflow: hidden;"> <div style="border: 0px solid transparent;color: #ec7211; display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.2; Margin: 25px 20px 25px 25px; max-height: none; max-width: none; padding: 0px; text-decoration: none;"> <div style=""> <font color="#000000" face="Arial, Helvetica, sans-serif" size="4"> <p><b>Informe académico de</b>: ' + student.toString().toUpperCase() + '</p> <p><b>Porcentaje de asistencia</b>: ' + attendance + ' %</p> <p><b>Porcentaje de KC realizados</b>: ' + kcok + ' %</p> <p><b>estado de los KC</b>: ' + state2 + '</p> <p><b>Cantidad de KC pendientes (sin hacer mas los que tienen baja nota)</b>: ' + totalfalt + '</p> <p><b>Cantidad de KC sin realizar</b>: ' + notDone + '</p> <p><b>Lista de los KC sin realizar</b>: </p> <ul>' + notDoneKcList + '</ul> <p><b>Cantidad de KC con baja nota</b>: ' + lowGrade + '</p> <p><b>Lista de KC con baja nota</b>:</p> <ul>' + lowGradeKcList + '</ul> <p>Cualquier inquietud, el grupo de docentes estamos para ayudarles.</p> <p>* Los KC con BAJA NOTA son aquiellos con menos del 70% de la nota maxima.</p> <p>* Recuerden realizar las "Notas of salida" (Exit Tickets).</p> </font> </div> </div> </td> </tr> </tbody> </table> <!--signature--> ' + signature + ' <!-- separador --> <table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr style="background-color: #ced2d5"> <td valign="top" class="stack-column" style="width: 100%; overflow: hidden;"> <div style="border: 0px solid transparent;color: rgb(0, 0, 0); display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 1px 25px; max-height: none; max-width: none; padding: 0px; text-decoration: none; height: 3px; background-color: #232f3e;"> </div> </td> </tr> </tbody> </table> <!-- footer --> <table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr style="background-color: #232f3e;"> <td valign="top" align="center" class="stack-column" style="width: 100%; overflow: hidden;"> <span class="eddie-wrapper" style="display: inline-block; Margin: 20px 10px 20px 0px"><a href="https://www.facebook.com/potrerodigital/" data-interests="" style="text-decoration: none;"><img src="http://cdn.pemres02.net/10433/recurso-9.png" alt="Facebook" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: auto; height: auto; display: inline-block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 0px; max-height: none; max-width: none; padding: 0px; text-decoration: none;" data-margin="20px 10px 20px 0px"></a></span> <span class="eddie-wrapper" style="display: inline-block; Margin: 20px 10px 20px 0px"><a href="https://www.linkedin.com/company/potrero-digital/mycompany/" data-interests="" style="text-decoration: none;"><img src="http://cdn.pemres02.net/10433/recurso-8.png" alt="LinkedIn" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: auto; height: auto; display: inline-block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 0px; max-height: none; max-width: none; padding: 0px; text-decoration: none;" data-margin="20px 10px 20px 0px"></a></span> <span class="eddie-wrapper" style="display: inline-block; Margin: 20px 10px 20px 0px"><a href="https://www.instagram.com/potrerodigital/" data-interests="" style="text-decoration: none;"><img src="http://cdn.pemres02.net/10433/recurso-10.png" alt="Instagram" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: auto; height: auto; display: inline-block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 0px; max-height: none; max-width: none; padding: 0px; text-decoration: none;" data-margin="20px 10px 20px 0px"></a></span> <span class="eddie-wrapper" style="display: inline-block; Margin: 20px 10px 20px 0px"><a href="https://www.youtube.com/channel/UCkh0OTzDBAtqKtXjHFqQinQ" data-interests="" style="text-decoration: none;"><img src="http://cdn.pemres02.net/10433/recurso-11.png" alt="Youtube" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: auto; height: auto; display: inline-block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 0px; max-height: none; max-width: none; padding: 0px; text-decoration: none;" data-margin="20px 10px 20px 0px"></a></span> </td> </tr> </tbody> </table> </div> </td> </tr> </tbody> </table> </td> </tr> </tbody> </table></body></html>');

      /* usar https://www.textfixer.com/html/compress-html-compression.php para 
generar html en una sola linea */


      var new_subject = "Informe Académico " + student.toString().toUpperCase() + " - AWS re/Start " + courseName;
      var empty_msj = "";

      /* send email to each student -*/
      GmailApp.sendEmail(email, new_subject, empty_msj, { cc: ccemail, attachments: body_html });
    })
  return (SpreadsheetApp.getActiveSpreadsheet().toast(
    'todos los informes academicos han sido enviados '
    , 'FINALIZADO el envio de e-mails'))
}

//-----------------------------------------------------------------------------
function renewingData() {
  //-----------------------------------------------------------------------------

  allProcess()
  preparSheets()
  updateAttendanceAndGrade()
}

//  -----------------------------------------------------------------------------
function allProcess() {
  //-----------------------------------------------------------------------------
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Inicio del procesamiento de los archivos .csv'
    , 'PROCESANDO...');
  while (csvFiles.hasNext()) {
    let csvFile = csvFiles.next()
    let csvFileId = csvFile.getId()
    let csvFileName = csvFile.getName()

    csvList[csvFileName] = csvFileId    // Object for csv File {name: id}
  }
  csvNames = Object.keys(csvList)   // Array for csv File names
  csvIds = Object.values(csvList)   // Array for csv File ids

  for (i = 0; i < csvNames.length; i++) {

    console.log("file name : " + csvNames[i])

    // Variables for csv Files of attendance

    if ((csvNames[i].match(attendanceRegex)) != null) {
      var spId = DriveApp.getFilesByName(attendanceSpName).next().getId()
      var regex = /(^.*-)(\d{8})(.*)/
      var newString = "$2"
      var encoding = 'UTF-16'
      var delimiter = '\t'
      var targetFolderName = attendanceBackupFolder

      // Variables for a csv Files of grade

    } else if ((csvNames[i].match(gradeRegex)) != null) {
      var spId = DriveApp.getFilesByName(gradeSpName).next().getId()
      var regex = /(.*)(\d{4})-(\d{2})-(\d{2}).*/
      var newString = "$2$3$4"
      var encoding = 'UTF-8'
      var delimiter = ','
      var targetFolderName = gradeBackupFolder

      // Variables for a csv Files of lab

    } else if ((csvNames[i].match(labRegex)) != null) {
      var spId = DriveApp.getFilesByName(labSpName).next().getId()
      var regex = /.*/
      var newString = systemDate
      var encoding = 'UTF-8'
      var delimiter = ','
      var targetFolderName = labBackupFolder

      // Variables for a other Files

    } else {
      var regex = "//"
      var newString = ""
      var spId = ""
      var targetFolderName = restBackupFolder
    }

    console.log("loop : " + i)

    var ssName = csvNames[i].replace(regex, newString);       // Sheet Names are replaced with regex names
    var csvAccess = DriveApp.getFileById(csvIds[i])
    var targetFolderAccess = DriveApp.getFoldersByName(targetFolderName)

    console.log("ssName with regex : " + ssName)

    if (spId != "") {

      var spAccess = SpreadsheetApp.openById(spId)

      if (!spAccess.getSheetByName(ssName)) {                 // Create a Sheet if she not exist
        spAccess.insertSheet(ssName);
      }

      var ssAccess = spAccess.getSheetByName(ssName)
      ssAccess.clearContents()

      // Load Data of csv File in to sheet
      var csvData = csvAccess.getBlob().getDataAsString(encoding).valueOf()
      var csv = Utilities.parseCsv(csvData, delimiter);
      var success = ssAccess.getRange(1, 1, csv.length, csv[0].length).setValues(csv);

      // If a load data is successly, moves csv file to backup folder
      if (success && targetFolderAccess.hasNext()) {
        csvAccess.moveTo(targetFolderAccess.next())
      }

    } else {
      csvAccess.moveTo(targetFolderAccess.next())
    }
  }
  return (SpreadsheetApp.getActiveSpreadsheet().toast(
    'Archivos procesados:\n' + csvNames.toString()
    , 'PROCESAMIENTO CONCLUIDO', 3)
  )
}

// Ordenamiento de las hojas de cada planilla
/*---------------------------------------------------------------------------*/
function preparSheets() {
  /*---------------------------------------------------------------------------*/

  for (i = 0; i < spNames.length; i++) {
    var spFullId = DriveApp.getFilesByName(spNames[i]).next().getId()
    var spFullAccess = SpreadsheetApp.openById(spFullId)
    var sss = spFullAccess.getSheets()
    if (!spFullAccess.setActiveSheet(spFullAccess.getSheetByName("Report"))) {
      var ssReport = sss[0].setName("Report")
    } else {
      var ssReport = spFullAccess.setActiveSheet(spFullAccess.getSheetByName("Report"));
    }
    sheetNames = []
    for (j = 0; j < sss.length; j++) {
      sheetNames.push(sss[j].getName());
    }
    sheetNames.sort().reverse();

    for (var k = 0; k < sheetNames.length; k++) {
      spFullAccess.setActiveSheet(spFullAccess.getSheetByName(sheetNames[k]));
      spFullAccess.moveActiveSheet(k + 1);
    }
    spFullAccess.setActiveSheet(spFullAccess.getSheetByName("Report"));
    spFullAccess.moveActiveSheet(1);

    // default data for Report sheet
    ssReport.getRange(1, 1, arrayStData.length, 4).setValues(arrayStData)
    // delete header of columns
    ssReport.getRange(1, 5, 1, ssReport.getLastColumn()).clearContent()
    var reportRangeHeaders = ssReport.getRange(1, 5, 1, sheetNames.length - 1)
    // complete header of columns with sheet names
    reportRangeHeaders.setValues([sheetNames.slice(1)])
  }
}

/*---------------------------------------------------------------------------*/
function updateAttendanceAndGrade() {
  /*---------------------------------------------------------------------------*/

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Inicio de la carga de nuevos datos'
    , 'ACTUALIZANDO esta planilla ...');

  const ssGradeLast = spGradeAccess.getSheets()[1]
  const arrayGradeData = ssGradeLast.getDataRange().getValues()

  var ssDataRange = []
  var ssDateName = []
  var course = {}
  // ---------------------------------------------------------------- Array with names of all ss 

  for (let h = 1; h <= (ssAll.length) - 1; h++) {
    ssDateName.push(ssAll[h].getName())
  }
  const reportRangeHeaders = ssReportAccess.getRange(1, 5, 1, ssDateName.length)
  reportRangeHeaders.setValues([ssDateName])

  console.log(reportRangeHeaders)

  for (var i = 0; i < ssDateName.length; i++) {

    var emails = {}
    var dte = ssDateName[i] //name of a sheet
    var ssAccess = spAttAccess.getSheetByName(dte) // Sheet access by your name
    var ssAccessReport = spAttAccess.getSheetByName('Report') // Sheet access by your name

    // ------------------------------------------------------ Array with data in everyone sheet

    var ssDataRange = ssAccess.getRange(2, 8, ssAccess.getLastRow() - 1, 3).getValues()

    for (let j = 0; j < ssDataRange.length; j++) {

      var email = ssDataRange[j][0]
      var sT = new Date(ssDataRange[j][1]) / 60000 // minutes
      var eT = new Date(ssDataRange[j][2]) / 60000 // minutes

      if (!emails[email]) {
        Object.assign(emails, { [email]: { [sT]: eT } })
      } else {
        Object.assign(emails[email], { [sT]: eT })
      }

    }

    Object.assign(course, { [dte]: emails })
    var reportEmails = ssAccessReport.getRange(2, 4, ssAccessReport.getLastRow() - 1, 1).getValues()

    var clearColumn = ssAccessReport.getRange(2, i + 1 + 4, ssAccessReport.getLastRow() - 1, 1).clearContent()


    for (let k = 0; k < reportEmails.length; k++) {

      var eMail = reportEmails[k][0]

      if (course[dte][eMail]) {
        var sTList = Object.keys(course[dte][eMail]).sort()
        ssAccessReport.getRange(k + 2, i + 1 + 4).setValue(1)
      } else {
        ssAccessReport.getRange(k + 2, i + 1 + 4).setValue(0)
      }

    }
  }
  spAttGradeAccess.getSheetByName('ASIST-WEBEX').getRange(1, 1, arrayStData.length, 4).setValues(arrayStData)
  spAttGradeAccess.getSheetByName('ASIST-WEBEX').getRange(1, 8, ssReportAccess.getLastRow(), ssReportAccess.getLastColumn() - 4).setValues(rangeDataRepor)

  spAttGradeAccess.getSheetByName('KC').getRange(1, 1, arrayStData.length, 4).setValues(arrayStData)

  ssDkc.clearContents()
  ssDkc.getRange(1, 1, ssGradeLast.getLastRow(), ssGradeLast.getLastColumn()).setValues(arrayGradeData)
  var ssGradeLastName = ssGradeLast.getName()
  ssKc.getRange(1, 2).setValue(ssGradeLastName)

  return (SpreadsheetApp.getActiveSpreadsheet().toast(
    'Planilla lista para compartir. ( Nota : recuerde agregar los encabezados'
    + ' de los KC mas recientes)'
    , 'ACTUALIZACION TERMINADA', 7))
}