//-------------------------------------------------------------------------------
function genesis() {
  //-----------------------------------------------------------------------------

  const courseNumber = 'TEST' // <<<------ CHANGE THE COURSE NUMBER HERE, BEFORE EXECUTING

  /*CREATION OF NECESSARY, FOLDERS AND FILES
  
      ARBUE_TEST/
      ├── FILE_INCOMING_TEST
      └── STORAGE_TEST
          ├── INPUT_TEST
          │   ├── ADMIN_AND_STUDENT_DATA_TEST
          │   │   ├── attendance_TEST.xlsx
          │   │   ├── grade_TEST.xlsx
          │   │   ├── GRADUATES_TEST
          │   │   ├── html_TEST.xlsx
          │   │   └── lab_TEST.xlsx
          │   ├── ATTENDANCE_BACKUP_TEST
          │   ├── GRADE_BACKUP_TEST
          │   ├── LAB_BACKUP_TEST
          │   └── REST_BACKUP_TEST
          ├── OUTPUT_TEST
          │   ├── ADMIN_TEST
          │   │   └── backoffice_TEST.xlsx
          │   ├── INSTRUCTOR_TEST
          │   └── STUDENT_TEST
          └── SCRIPT_TEST

          15 directories, 5 files

*/

  // FOLDER'S LAYERS
  const layer_5 = ['ADMIN_AND_STUDENT_DATA_', ['GRADUATES_']]
  const layer_4 = ['INPUT_', ['ADMIN_AND_STUDENT_DATA_', 'ATTENDANCE_BACKUP_', 'GRADE_BACKUP_',
    'LAB_BACKUP_', 'REST_BACKUP_']]
  const layer_3 = ['OUTPUT_', ['ADMIN_', 'STUDENT_', 'INSTRUCTOR_']]
  const layer_2 = ['STORAGE_', ['SCRIPT_', 'OUTPUT_', 'INPUT_']]
  const layer_1 = ['ARBUE_', ['FILE_INCOMING_', 'STORAGE_']]

  const layers = [layer_1, layer_2, layer_3, layer_4, layer_5]

  var fIds = {}

  // FOLDER NAME + COURSE NUMBER
  for (var i = 0; i < layers.length; i++) {
    layers[i][0] = layers[i][0] + courseNumber
    for (let j = 0; j < layers[i][1].length; j++) {
      layers[i][1][j] = layers[i][1][j] + courseNumber
    }
  }
  // FOLDER CREATION
  for (var i = 0; i < layers.length; i++) {

    if (!fIds[layers[i][0]]) {

      var fatherFolder = DriveApp.createFolder(layers[i][0])
      var fatherFolderId = fatherFolder.getId()
      var fatherFolderName = fatherFolder.getName()
      var fatherFolderAccess = DriveApp.getFolderById(fatherFolderId)

      Object.assign(fIds, { [fatherFolderName]: fatherFolderId })

      console.log("father folder :" + fatherFolderName + "successfully created")

    } else {
      var fatherFolderAccess = DriveApp.getFolderById(fIds[layers[i][0]])
    }

    for (var j = 0; j < layers[i][1].length; j++) {

      var childFolder = fatherFolderAccess.createFolder(layers[i][1][j])
      var childFolderName = childFolder.getName()
      var childFolderId = childFolder.getId()
      Object.assign(fIds, { [childFolderName]: childFolderId })

      console.log("child folder :" + childFolderName + "successfully created")

    }
  }

  //CREATION OF NECESSARY SPREADSHEETS 
  // Names of Spreadsheets
  const attendanceSpName = 'attendance_' + courseNumber
  const gradeSpName = 'grade_' + courseNumber
  const labSpName = 'lab_' + courseNumber
  //const studentSpName = 'studentData_' + courseNumber
  const backofficeSpName = 'backoffice_' + courseNumber
  const htmlSpName = 'html_' + courseNumber

  var fnameList = ['ADMIN_AND_STUDENT_DATA_', 'ADMIN_AND_STUDENT_DATA_', 'ADMIN_AND_STUDENT_DATA_', 'ADMIN_', 'ADMIN_AND_STUDENT_DATA_']
  var spNamesList = [attendanceSpName, gradeSpName, labSpName, backofficeSpName, htmlSpName]
  //var spNamesList = [attendanceSpName, gradeSpName, labSpName, studentSpName, backofficeSpName, htmlSpName]

  var sPIds = {}

  for (let i = 0; i < spNamesList.length; i++) {
    var spCreate = SpreadsheetApp.create(spNamesList[i])
    var spId = spCreate.getId()
    var spNam = spCreate.getName()

    DriveApp.getFileById(spId).moveTo(DriveApp.getFolderById(fIds[fnameList[i] + courseNumber]))

    console.log(" Spreadsheet :" + spNam + "successfully created in to " + fIds[fnameList[i] + courseNumber] + " folder")

    Object.assign(sPIds, { [spNam]: spId })
  }

  // Spreadsheet htmlSpName
  var spHtmlAccess = SpreadsheetApp.openById(sPIds[htmlSpName])
  spHtmlAccess.getSheets()[0].setName('BODY')
  spHtmlAccess.insertSheet('SIGNATURE')
  spHtmlAccess.insertSheet('INSTRUCTOR_DATA')

  // HTML CODE FOR EMAIL'S BODY AND SIGNATURE
  const body = ['INFORME ACADEMICO', '<html><head> <meta charset="utf-8"> <meta name="viewport" content="width=device-width"> <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"> <meta http-equiv="X-UA-Compatible" content="IE=edge"> <meta name="robots" content="noindex"> <base target="_blank"> <style type="text/css"> body, div[style*="margin: 16px 0"], html { margin: 0 !important } body, html { padding: 0 !important; height: 100% !important; width: 100% !important } * { -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100% } table, td { mso-table-lspace: 0 !important; mso-table-rspace: 0 !important } table { border-spacing: 0 !important; border-collapse: collapse !important; margin: 0 auto !important } table table table { table-layout: auto } img { -ms-interpolation-mode: bicubic } .yshortcuts a { border-bottom: none !important } .mobile-link--footer a, a[x-apple-data-detectors] { color: inherit !important; text-decoration: underline !important } .firma{ background-color: #232f3e; } @media screen and (max-width:600px) { .stack-column-half { width: 50% !important; display: inline-block !important } .center-on-narrow, .fluid, .fluid-centered { margin-left: auto !important; margin-right: auto !important } table { table-layout: fixed !important } .email-container { width: 100% !important } .fluid, .fluid-centered { max-width: 100% !important; height: auto !important } .stack-column, .stack-column-center, .stack-column-full-width { display: block !important; width: 100% !important; max-width: 100% !important; direction: ltr !important } .center-on-narrow { display: block !important; float: none !important } table.center-on-narrow { display: inline-block !important } .stack-column-full-width .eddie-wrapper { color: white; width: 100% } } </style></head><body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" offset="0"> <table cellspacing="0" cellpadding="0" border="0" width="100%" style="font-family: Helvetica, Arial, sans-serif; width: 100%; padding: 20px; background-color: rgb(235, 235, 235); background-image: none;"> <tbody> <tr> <td align="center"> <table class="email-container" width="660"> <tbody> <tr> <td> <div class="eddie-page"> <!-- barra inicial --> <table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr style="background-color: #232f3e;"> <td valign="top" align="center" stackclass="stack-column-full-width" class="stack-column-full-width" style="width: 50%; overflow: hidden;"> <img src="https://d1.awsstatic.com/training-and-certification/Logos/aws_restart_logo_reverse.860113148166c4742ebd63e8fa74d09ae4cf64ea.png" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width:50%; color: rgb(0, 0, 0); display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 5px; max-height: none; max-width: none; padding: 5px; text-decoration: none; min-height: 10px;" alt=""> </td> <td valign="top" align="left" stackclass="stack-column-full-width" class="stack-column-full-width" style="width: 50%; overflow: hidden;"> <img src="https://static.wixstatic.com/media/5b90eb_2f1f983af79a4e69ba942bc0586dbb7d~mv2.png/v1/fill/w_382,h_40,al_c,q_85,usm_0.66_1.00_0.01,enc_auto/potrero_digital_2021_edited.png" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: 80%; color: rgb(0, 0, 0); display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 10px; max-height: none; max-width: none; padding: 10px; text-decoration: none; min-height: 10px;" alt=""> </td> </tr> </tbody> </table> <!-- imagen grande --> <table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr style="background-color: #ec7211;"> <td valign="top" class="center-on-narrow stack-column" style="width: 100%; overflow: hidden;"> <div style="border: 0px solid transparent;color: #ec7211; display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.2; Margin: 25px 20px 25px 25px; max-height: none; max-width: none; padding: 0px; text-decoration: none;"> <div style="text-align:center;"> <font color="white" face="Arial, Helvetica, sans-serif" size="7"> <b>Informe Académico</b> </font> </div> </div> </td> </tr> </tbody> </table> <!-- cuerpo --> <table cellspacing="0" cellpadding="0" border="0" style="width: 700px; height: 528px;"> <tbody> <tr style="background-color: #ced2d5;"> <td valign="top" class="center-on-narrow stack-column" style="width: 100%; overflow: hidden;"> <div style="border: 0px solid transparent;color: #ec7211; display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.2; Margin: 25px 20px 25px 25px; max-height: none; max-width: none; padding: 0px; text-decoration: none;"> <div style=""> <font color="#000000" face="Arial, Helvetica, sans-serif" size="4"> <p><b>Informe académico de</b>: {{student}} </p> <p><b>Porcentaje de asistencia</b>: {{attendance}} %</p> <p><b>Porcentaje de KC realizados</b>: {{kcok}} %</p> <p><b>Estado de los KC</b>: {{state2}}</p> <p><b>Cantidad de KC pendientes ("sin realizar" + "baja nota")</b>: {{totalfalt}}</p> <p><b>Cantidad de KC sin realizar</b>: {{notDone}}</p> <p><b>Lista de los KC sin realizar</b>: </p> <ul>{{notDoneKcList}}</ul> <p><b>Cantidad de KC con baja nota</b>: {{lowGrade}}</p> <p><b>Lista de KC con baja nota</b>:</p> <ul>{{lowGradeKcList}}</ul> <p>Cualquier inquietud, el grupo de docentes estamos para ayudarles.</p> <p>* Los KC con BAJA NOTA son aquiellos con menos del 70% de la nota maxima.</p> <p>* Recuerden realizar las "Notas de salida" (Exit Tickets).</p> </font> </div> </div> </td> </tr> </tbody> </table> <!--Firma--> {{signature}} <!-- separador --> <table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr style="background-color: #ced2d5"> <td valign="top" class="stack-column" style="width: 100%; overflow: hidden;"> <div style="border: 0px solid transparent;color: rgb(0, 0, 0); display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 1px 25px; max-height: none; max-width: none; padding: 0px; text-decoration: none; height: 3px; background-color: #232f3e;"> </div> </td> </tr> </tbody> </table> <!-- footer --> <table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr style="background-color: #232f3e;"> <td valign="top" align="center" class="stack-column" style="width: 100%; overflow: hidden;"> <span class="eddie-wrapper" style="display: inline-block; Margin: 20px 10px 20px 0px"><a href="https://www.facebook.com/potrerodigital/" data-interests="" style="text-decoration: none;"><img src="http://cdn.pemres02.net/10433/recurso-9.png" alt="Facebook" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: auto; height: auto; display: inline-block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 0px; max-height: none; max-width: none; padding: 0px; text-decoration: none;" data-margin="20px 10px 20px 0px"></a></span> <span class="eddie-wrapper" style="display: inline-block; Margin: 20px 10px 20px 0px"><a href="https://www.linkedin.com/company/potrero-digital/mycompany/" data-interests="" style="text-decoration: none;"><img src="http://cdn.pemres02.net/10433/recurso-8.png" alt="LinkedIn" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: auto; height: auto; display: inline-block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 0px; max-height: none; max-width: none; padding: 0px; text-decoration: none;" data-margin="20px 10px 20px 0px"></a></span> <span class="eddie-wrapper" style="display: inline-block; Margin: 20px 10px 20px 0px"><a href="https://www.instagram.com/potrerodigital/" data-interests="" style="text-decoration: none;"><img src="http://cdn.pemres02.net/10433/recurso-10.png" alt="Instagram" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: auto; height: auto; display: inline-block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 0px; max-height: none; max-width: none; padding: 0px; text-decoration: none;" data-margin="20px 10px 20px 0px"></a></span> <span class="eddie-wrapper" style="display: inline-block; Margin: 20px 10px 20px 0px"><a href="https://www.youtube.com/channel/UCkh0OTzDBAtqKtXjHFqQinQ" data-interests="" style="text-decoration: none;"><img src="http://cdn.pemres02.net/10433/recurso-11.png" alt="Youtube" width="auto" height="auto" class="fluid" style="border: 0px solid transparent;width: auto; height: auto; display: inline-block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.35; Margin: 0px; max-height: none; max-width: none; padding: 0px; text-decoration: none;" data-margin="20px 10px 20px 0px"></a></span> </td> </tr> </tbody> </table> </div> </td> </tr> </tbody> </table> </td> </tr> </tbody> </table></body></html>']
  const signature = ['AWS reStart', '<table cellspacing="0" cellpadding="0" border="0" style="width: 100%;"> <tbody> <tr class="firma" style=""> <td valign="top" align="center" class="center-on-narrow stack-column" style="width: 100%; overflow: hidden;"> <div style="border: 0px solid transparent;color: #f1f4f6; display: block; font-family: Helvetica, Arial, sans-serif; font-size: 10px; font-weight: normal; line-height: 1.2; Margin: 25px 20px 25px 25px; max-height: none; max-width: none; padding: 0px; text-decoration: none;"> <div> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="4" class="nombre"> <b>{{instructorName[email_prof]}}</b> | {{instructorRole[email_prof]}} </font> </div> <div> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="3">AWS Certified Associate<br></font> <div class=""> <font size="3"><b> <font face="&quot;Arial&quot;, Helvetica, sans-serif"> <a href="http://compromiso.org/" style="color: #ec7211;" class="">compromiso.org</a> </font> </b></font> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="3">| </font> <font size="3"><b> <font face="&quot;Arial&quot;, Helvetica, sans-serif"> <a href="http://potrerodigital.org/" class="" style="color: #ec7211;">potrerodigital.org</a> </font> </b></font> <br> </div> <div class="firm"> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="3"> Cel. +{{instructorCel[email_prof]}}</font> <font color="#f1f4f6" face="Arial, Helvetica, sans-serif" size="3">| </font> <font size="2"><b> <font face="&quot;Arial&quot;, Helvetica, sans-serif"> <a href="{{instructorLinkedin[email_prof]}}" class="" style="color: #ec7211;">LinkedIn</a> </font> </b></font> <br> </div> </div> </div> </td> </tr> <tr class="firma"> <td valign="top" align="center" class="center-on-narrow stack-column" style="width: 100%; overflow: hidden; vertical-align: bottom; padding-bottom: 15px;"> <img src="https://ci3.googleusercontent.com/proxy/ep9w9gvBrTwl4kJ19bQf0B3BaFV-O9Bfd1ooVx2pJWVong0E4Sxa2NdRdZ5Atfs36cC13_SfW3IeTqCmO9lneyty4VFwiSvX7QC096MWg_sDU0o8-EvZBX1FWgbS5FH3q1DWlDqudPGTeUEpNpiM=s0-d-e1-ft#https://images.credly.com/size/340x340/images/44e2c252-5d19-4574-9646-005f7225bf53/image.png" alt="AWS re/Start Graduate" width="96" height="96" class="CToWUd" data-bit="iit"><img src="https://ci4.googleusercontent.com/proxy/ZDZm9x8NP-OYdpemo9sXn8iq8qc7K4FslWZBBTo1zg21pCjv13Ph6KMzbIU0g2oAkEU71HA-gMsWoqGwPYj7vWpnD7xZARpNotH0BnphC7aMccAc7618dQO8o_dLzeDBDGu2yiynJvrX1or5KzfM=s0-d-e1-ft#https://images.credly.com/size/340x340/images/00634f82-b07f-4bbd-a6bb-53de397fc3a6/image.png" alt="AWS Certified Cloud Practitioner" width="96" height="96" class="CToWUd" data-bit="iit"> <img src="https://ci3.googleusercontent.com/proxy/VpBzkJjDHaXcQzbidXDpRdPSPahgJpbU0hDNcIlrLBYVjLsxUevGLzYdvNk6KUL8LCGpejdGVl6U02zuSjo3ga91sTAokXYo1Tm1Z3t0Kc1p6h4xg6shcgYuMVmy_rP_SmFkXtreax1p1qLovE3N=s0-d-e1-ft#https://images.credly.com/size/340x340/images/0e284c3f-5164-4b21-8660-0d84737941bc/image.png" alt="AWS Certified Solutions Architect – Associate" width="96" height="96" class="CToWUd" data-bit="iit"> <img src="https://ci6.googleusercontent.com/proxy/R-aIAR8Uu3ffYUJisbl6bfL1YWUP8r_JXip-EIbcQV-er5fz-wbsvK4HA7qItSyDrVV70pze-9hgT6MPwt2pfkYAIUtPwNtIASl_1MxA1MCaVn4soT1EGZpv5XsakDHw8PqTsVXIvbFtVLfQaZqw=s0-d-e1-ft#https://images.credly.com/size/340x340/images/e426d40e-8a6a-4f72-866e-2abfcfbde46b/image.png" alt="AWS re/Start Accredited Instructor" width="96" height="96" class="CToWUd" data-bit="iit"> </td> </tr> </tbody></table>']

  const instructorData = [['EMAIL', 'NAME', 'LINKEDIN', 'CELULAR', 'ROLE'],
  ['leandro.garrido@compromiso.org',
    'Leandro Garrido',
    'https://www.linkedin.com/in/leandro-garrido/',
    '54 9 11 5895-3808',
    'Instructor'],
  ['ariel.orange@compromiso.org',
    'Ariel Orange',
    'https://www.linkedin.com/in/ariel-orange/',
    '54 9 11 3388-1887',
    'Instructor'],
  ['emiliano.piai@compromiso.org',
    'Emiliano Piai',
    'https://www.linkedin.com/in/emiliano-piai-826b7a233/',
    '54 9 2616 65-1022',
    'Tutor']]
  spHtmlAccess.getSheetByName('BODY').getRange(1, 1, 1, 2).setValues([body])
  spHtmlAccess.getSheetByName('SIGNATURE').getRange(1, 1, 1, 2).setValues([signature])
  spHtmlAccess.getSheetByName('INSTRUCTOR_DATA').getRange(1, 1, 4, 5).setValues(instructorData)

  // HIDEN SHEET BODY AND SIGNATURE 
  spHtmlAccess.getSheetByName('BODY').hideSheet()
  spHtmlAccess.getSheetByName('SIGNATURE').hideSheet()

  // Spreadsheet "attendanceSpName","gradeSpName" and  "labSpName"
  for (let i = 0; i < 5; i++) {
    var spAcc = SpreadsheetApp.openById(sPIds[spNamesList[i]])
    var ssReport = spAcc.getSheets()[0]
    ssReport.setName('Report')
    if (i == 3) {
      spAcc.insertSheet('kc_asist')
    } else {
      ssReport.getRange(1, 1, 2, 4).setValues([['family name', 'first name', 'ID', 'SIS Login ID'], ['Estudiante', 'de prueba', '', '']])
    }

  }
}
