//-------------------------------------------------------------------------------
function genesis() {
  //-----------------------------------------------------------------------------

  const courseNumber = 'TEST' // <<<------ CAMBIAR AQUI EL NUMERO DE CURSO ANTES DE EJECUTAR

  const layer_5 = ['ADMIN_AND_STUDENT_DATA_', ['GRADUATES_']]
  const layer_4 = ['INPUT_', ['ADMIN_AND_STUDENT_DATA_', 'ATTENDANCE_BACKUP_', 'GRADE_BACKUP_',
    'LAB_BACKUP_', 'REST_BACKUP_']]
  const layer_3 = ['OUTPUT_', ['ADMIN_', 'STUDENT_', 'INSTRUCTOR_']]
  const layer_2 = ['STORAGE_', ['SCRIPT_', 'OUTPUT_', 'INPUT_']]
  const layer_1 = ['ARBUE_', ['FILE_INCOMING_', 'STORAGE_']]

  const layers = [layer_1, layer_2, layer_3, layer_4, layer_5]
  var fIds = {}

  for (var i = 0; i < layers.length; i++) {
    layers[i][0] = layers[i][0] + courseNumber
    for (let j = 0; j < layers[i][1].length; j++) {
      layers[i][1][j] = layers[i][1][j] + courseNumber
    }
  }

  for (var i = 0; i < layers.length; i++) {

    if (!fIds[layers[i][0]]) {

      var fatherFolder = DriveApp.createFolder(layers[i][0])
      var fatherFolderId = fatherFolder.getId()
      var fatherFolderName = fatherFolder.getName()
      var fatherFolderAccess = DriveApp.getFolderById(fatherFolderId)

      Object.assign(fIds, { [fatherFolderName]: fatherFolderId })

    } else {
      var fatherFolderAccess = DriveApp.getFolderById(fIds[layers[i][0]])
    }

    for (var j = 0; j < layers[i][1].length; j++) {

      var newFolder = fatherFolderAccess.createFolder(layers[i][1][j])
      var newFolderName = newFolder.getName()
      var newFolderId = newFolder.getId()
      Object.assign(fIds, { [newFolderName]: newFolderId })
    }
  }
}
