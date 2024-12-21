function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Run')
    .addItem('Sync', 'sync')
    .addItem('Update External Share Expiry', 'updateExternalShareExpiry')
    .addToUi();
}

function updateExternalShareExpiry() {
  const configData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config').getDataRange().getValues();
  const rootFolderId = getConfigValue(configData, 'Data root folder');
  const expiryTime = getConfigValue(configData, "Expiry Datetime")
  const allFiles = recursiveDriveList(rootFolderId)
  allFiles.forEach(f => {
    const allPermissions = Drive.Permissions.list(f.id, { supportsAllDrives: true, fields: "*" })
    const externalPermissionsNeedsUpdating = allPermissions.permissions
      .filter(p => !p.emailAddress.includes("woodcraft.org.uk"))
      .filter(p => p.expirationTime !== expiryTime)
    externalPermissionsNeedsUpdating.forEach(p => {
      Drive.Permissions.update({ role: p.role, expirationTime: expiryTime }, f.id, p.id, { supportsAllDrives: true })
    })
  })
}


function sync() {
  // Load some data
  const configData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config').getDataRange().getValues();
  const rootFolderId = getConfigValue(configData, 'Data root folder');
  const rootFolder = DriveApp.getFolderById(rootFolderId)

  const templateData = transpose(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Templates').getDataRange().getValues())

  const districtData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Districts').getDataRange().getValues();
  const groupsData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Groups').getDataRange().getValues();

  const year = new Date().getFullYear()

  //no headers
  districtData.shift()
  groupsData.shift()
  templateData[0].shift() //district copy
  templateData[1].shift() //district shortcut
  templateData[2].shift() //group copy
  templateData[3].shift() //group shortcut

  // Check if all groups have a district
  groupsData.forEach(group => {
    const district = districtData.find(d => d[0] === group[1]);
    if (!district) {
      console.log('District not found for group', group)
      throw new Error('District not found for group')
    }
  })

  // Create district folders

  const districtFolders = {}

  districtData.forEach(district => {

    const districtFolderName = applyTemplate(getConfigValue(configData, "District Folder Name Template"), [['%D', district[0]], ['%Y', year]])
    const districtFolder = createOrGetFolder(districtFolderName, rootFolder)
    districtFolders[district[0]] = districtFolder

    const editors = districtFolder.getEditors()

    if (editors.find(e => e.getEmail() === district[1])) {
      console.log('District folder already shared with', district[1])
    } else {
      districtFolder.addEditor(district[1])
      console.log('Sharing District folder with', district[1])
    }

    console.log('District folder created, copying templates', districtFolder.getId())

    templateData[0].forEach(template => {
      const file = DriveApp.getFileById(template)
      const newName = applyTemplate(getConfigValue(configData, "District File Name Template"), [['%F', file.getName()], ['%D', district[0]], ['%Y', year]])

      createOrCopyFile(file, newName, districtFolder)
    })

    templateData[1].forEach(template => {
      const file = DriveApp.getFileById(template)
      createShortcutIfMissing(file, districtFolder)
    })

    console.log('Templates copied')
  })

  // Create group folders
  groupsData.forEach(group => {
    const districtFolder = districtFolders[group[1]]
    const groupFolderName = applyTemplate(getConfigValue(configData, "Group Folder Name Template"), [['%G', group[0]], ['%D', group[1]], ['%Y', year]])
    const groupFolder = createOrGetFolder(groupFolderName, districtFolder)

    const editors = groupFolder.getEditors()

    if (editors.find(e => e.getEmail() === group[2])) {
      console.log('Group folder already shared with', group[2])
    } else {
      groupFolder.addEditor(group[2])
      console.log('Sharing Group folder with', group[2])
    }

    console.log('Group folder created, copying templates', groupFolder.getId())

    templateData[2].forEach(template => {
      const file = DriveApp.getFileById(template)
      const newName = applyTemplate(getConfigValue(configData, "Group File Name Template"), [['%F', file.getName()], ['%G', group[0]], ['%D', group[1]], ['%Y', year]])

      createOrCopyFile(file, newName, groupFolder)
    })

    templateData[3].forEach(template => {
      const file = DriveApp.getFileById(template)
      createShortcutIfMissing(file, groupFolder)
    })

    console.log('Templates copied')
  })
}

function createOrGetFolder(name, parent) {
  const query = `title = '${name}' and '${parent.getId()}' in parents and trashed = false`
  const folderSearch = parent.searchFolders(query);

  if (folderSearch.hasNext()) {
    console.log('Folder already exists', name)
    return folderSearch.next()
  } else {
    console.log('Creating folder', name)
    return parent.createFolder(name)
  }
}

function createOrCopyFile(file, newName, parent) {
  const query = `title = '${newName}' and '${parent.getId()}' in parents and trashed = false`
  const fileSearch = parent.searchFiles(query);

  if (fileSearch.hasNext()) {
    console.log('File already exists', newName)
    return fileSearch.next()
  } else {
    console.log('Creating file', newName)
    file.makeCopy(newName, parent)
  }
}

function createShortcutIfMissing(file, parent) {
  const query = `title = '${file.getName()}' and '${parent.getId()}' in parents and trashed = false`
  const templateFileSearch = parent.searchFiles(query);

  if (templateFileSearch.hasNext()) {
    console.log('Shortcut already exists for ', file.getName())
  } else {
    console.log('Creating shortcut', file.getName())
    parent.createShortcut(file.getId())
  }
}

function applyTemplate(template, substitutions) {
  return substitutions.reduce((acc, sub) => acc.replace(sub[0], sub[1]), template)
}

function getConfigValue(configData, item) {
  return configData.find(i => i[0] === item)[1]
}

function recursiveDriveList(root) {
  const list = Drive.Files.list({ q: `'${root}' in parents and trashed = false`, includeTeamDriveItems: true, supportsAllDrives: true, fields: "files(id, name, mimeType, permissionIds)" })
  let results = list.files.filter(f => f.mimeType !== "application/vnd.google-apps.folder" && f.mimeType !== "application/vnd.google-apps.shortcut")
  const folders = list.files.filter(f => f.mimeType == "application/vnd.google-apps.folder")
  folders.forEach(f => {
    results = [...results, ...recursiveDriveList(f.id)]
  })
  return results
}

const transpose = arr => arr.reduce((m, r) => (r.forEach((v, i) => (m[i] = (m[i] || []), m[i].push(v))), m), [])