const systemDate = Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd")

const numCourse = 'TEST'

const attendanceSpName = 'attendance_' + numCourse
const gradeSpName = 'grade_' + numCourse
const labSpName = 'lab_' + numCourse

const sourceFolderName = 'FILE_INCOMING_' + numCourse

const attendanceBackupFolder = 'ATTENDANCE_BACKUP_' + numCourse
const gradeBackupFolder = 'GRADE__BACKUP_' + numCourse
const labBackupFolder = 'LAB_BACKUP_' + numCourse

function attendanceProcess() { allProcess('attendance')}
function gradeProcess() { allProcess(grade)}
function labProcess() { allProcess(lab)}

function allProcess(area) {
  switch (area) {
    case 'attendance':
      var targetFolderName = attendanceBackupFolder
      break;
    case 'grade':
      var targetFolderName = gradeBackupFolder
    case 'lab':
      var targetFolderName = labBackupFolder
      break;
    default:
      console.log(`Sorry, we are out of ${area}.`);
  }

  var csvFolder = DriveApp.getFoldersByName(sourceFolderName).next()
  var csvFiles = csvFolder.searchFiles('modifiedDate > "' + systemDate + '"')
  
  while (csvFiles.hasNext()) {
    var csvFile = csvFiles.next()
    var csvFileId = csvFile.getId()
    var csvFileType = csvFile.getMimeType()
    var csvFileData = csvFile.getDateCreated()
    console.log(csvFile.getName()+" "+csvFileType+" "+csvFileData)

  if (csvFileData > systemDate) {var csvMove = csvFile.moveTo(DriveApp.getFoldersByName(targetFolderName).next())}

      }
  

}