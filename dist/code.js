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

  const templateData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Templates').getDataRange().getValues()

  const districtData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Districts').getDataRange().getValues();
  const groupsData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Groups').getDataRange().getValues();

  const year = new Date().getFullYear()

  //no headers
  districtData.shift()
  groupsData.shift()
  templateData.shift() //district copy

  // Check if all groups have a district
  groupsData.forEach(group => {
    const district = districtData.find(d => d[0] === group[2]);
    if (!district) {
      console.log('District not found for group', group)
      throw new Error('District not found for group')
    }
  })

  // Create district folders

  const districtFolders = {}

  districtData.forEach(district => {

    const name = district[0]
    const number = district[1]
    const email = district[2]

    const districtFolderName = applyTemplate(getConfigValue(configData, "District Folder Name Template"), [['%D', name], ['%N', number], ['%Y', year]])
    const districtFolder = createOrGetFolder(districtFolderName, rootFolder)
    districtFolders[name] = districtFolder

    const editors = districtFolder.getEditors()

    if (editors.find(e => e.getEmail() === email)) {
      console.log('District folder already shared with', email)
    } else {
      districtFolder.addEditor(email)
      console.log('Sharing District folder with', email)
    }

    console.log('District folder created, copying templates', districtFolder.getId())

    templateData.filter(r => r[1] === "district" && r[2] === "copy").forEach(template => {
      const file = DriveApp.getFileById(getIdFromUrl(template[0]))
      const newName = applyTemplate(template[3], [['%F', file.getName()], ['%D', name], ['%N', number], ['%Y', year]])

      createOrCopyFile(file, newName, districtFolder)
    })

    templateData.filter(r => r[1] === "district" && r[2] === "shortcut").forEach(template => {
      const file = DriveApp.getFileById(getIdFromUrl(template[0]))
      createShortcutIfMissing(file, districtFolder)
    })

    console.log('Templates copied')
  })

  // Create group folders
  groupsData.forEach(group => {
    const name = group[0]
    const number = group[1]
    const district = group[2]
    const email = group[3]
    const districtFolder = districtFolders[district]
    const groupFolderName = applyTemplate(getConfigValue(configData, "Group Folder Name Template"), [['%G', name], ['%N', number], ['%D', district], ['%Y', year]])
    const groupFolder = createOrGetFolder(groupFolderName, districtFolder)

    const editors = groupFolder.getEditors()

    if (editors.find(e => e.getEmail() === email)) {
      console.log('Group folder already shared with', email)
    } else {
      groupFolder.addEditor(email)
      console.log('Sharing Group folder with', email)
    }

    console.log('Group folder created, copying templates', groupFolder.getId())

    templateData.filter(r => r[1] === "group" && r[2] === "copy").forEach(template => {
      const file = DriveApp.getFileById(getIdFromUrl(template[0]))
      const newName = applyTemplate(template[3], [['%F', file.getName()], ['%G', name], ['%N', number], ['%D', district], ['%Y', year]])

      createOrCopyFile(file, newName, groupFolder)
    })

    templateData.filter(r => r[1] === "group" && r[2] === "shortcut").forEach(template => {
      const file = DriveApp.getFileById(getIdFromUrl(template[0]))
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

function getIdFromUrl(url) {
  const pattern = /\/d\/([^\/\\]+)(?:\/|$)/i;
  return pattern.test(url || '')
    ? url.match(pattern)[1]
    : null;
}