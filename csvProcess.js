const numCourse = 'TEST'
const systemDate = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd")

const sourceFolderName = 'FILE_INCOMING_' + numCourse
const spFolderName = 'INPUT_' + numCourse

const attendanceBackupFolder = 'ATTENDANCE_BACKUP_' + numCourse
const gradeBackupFolder = 'GRADE_BACKUP_' + numCourse
const labBackupFolder = 'LAB_BACKUP_' + numCourse
const restBackupFolder = 'REST_BACKUP_' + numCourse

const attendanceSpName = 'attendance_' + numCourse

const gradeSpName = 'grade_' + numCourse
const labSpName = 'lab_' + numCourse

allProcess()

//-----------------------------------------------------------------------------
function run() { }
//-----------------------------------------------------------------------------
function allProcess() {
  //-----------------------------------------------------------------------------

  var csvFolderAccess = DriveApp.getFoldersByName(sourceFolderName).next()
  var spFolderAccess = DriveApp.getFoldersByName(spFolderName).next()
  var csvList = {}

  var csvFiles = csvFolderAccess.getFiles()

  while (csvFiles.hasNext()) {
    var csvFile = csvFiles.next()

    var csvFileId = csvFile.getId()
    var csvFileName = csvFile.getName()

    csvList[csvFileName] = csvFileId
  }

  var csvNames = Object.keys(csvList)
  var csvIds = Object.values(csvList)

  Logger.log("csvList --> " + csvNames + " " + csvIds)

  //example fo csv name = /participant-20230502222213.csv/
  var attendanceRegex = /participant/
  //example fo csv name = /2023-05-04T1109_Calificaciones-AWS_ARBUE13.csv/
  var gradeRegex = /Calificaciones/
  //example fo csv name = /RESTCUR-000001 ES 23732-1767-a0R4N00000G9qOGUAZ_Labtime_1683209549930.csv/
  var labRegex = /Labtime/

  for (i = 0; i < csvNames.length; i++) {

    console.log("file name : " + csvNames[i])

    if ((csvNames[i].match(attendanceRegex)) != null) {
      var spId = DriveApp.getFilesByName(attendanceSpName).next().getId()
      var regex = /(^.*-)(\d{8})(.*)/
      var newString = "$2"
      var encoding = 'UTF-16'
      var delimiter = '\t'
      var targetFolderName = attendanceBackupFolder

    } else if ((csvNames[i].match(gradeRegex)) != null) {
      var spId = DriveApp.getFilesByName(gradeSpName).next().getId()
      var regex = /(^\d+)-(\d+)-(\d+).*/
      var newString = "$1$2$3"
      var encoding = 'UTF-8'
      var delimiter = ','
      var targetFolderName = gradeBackupFolder

    } else if ((csvNames[i].match(labRegex)) != null) {
      var spId = DriveApp.getFilesByName(labSpName).next().getId()
      var regex = /.*/
      var newString = systemDate
      var encoding = 'UTF-8'
      var delimiter = ','
      var targetFolderName = labBackupFolder

    } else {
      var regex = "//"
      var newString = ""
      var spId = ""
      var targetFolderName = restBackupFolder
    }

    console.log("loop : " + i)

    var ssName = csvNames[i].replace(regex, newString);
    var csvAccess = DriveApp.getFileById(csvIds[i])
    var targetFolder = DriveApp.getFoldersByName(targetFolderName)

    console.log("ssName with regex : " + ssName)

    if (spId != "") {

      var spAccess = SpreadsheetApp.openById(spId)

      if (!spAccess.getSheetByName(ssName)) {
        spAccess.insertSheet(ssName);
      }

      var ssAccess = spAccess.getSheetByName(ssName)

      var csvData = csvAccess.getBlob().getDataAsString(encoding).valueOf()
      var csv = Utilities.parseCsv(csvData, delimiter);

      ssAccess.clearContents()
      var success = ssAccess.getRange(1, 1, csv.length, csv[0].length).setValues(csv);

      if (success | targetFolder.hasNext()) { csvAccess.moveTo(targetFolder.next()) }

    }else { csvAccess.moveTo(targetFolder.next()) }
  }
}