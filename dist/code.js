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
    const allPermissions = Drive.Permissions.list(f.id, { supportsAllDrives: true, includeItemsFromAllDrives: true, fields: "*" })
    const externalPermissionsNeedsUpdating = allPermissions.permissions
      .filter(p => !p.emailAddress.includes("woodcraft.org.uk"))
      .filter(p => p.expirationTime !== expiryTime)
    externalPermissionsNeedsUpdating.forEach(p => {
      Drive.Permissions.update({ role: p.role, expirationTime: expiryTime }, f.id, p.id, { supportsAllDrives: true, includeItemsFromAllDrives: true })
    })
  })
}


function sync() {
  // Load some data
  const configData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config').getDataRange().getValues();
  const rootFolderId = getConfigValue(configData, 'Data root folder');
  const rootFolder = Drive.Files.get(rootFolderId, {supportsAllDrives: true, includeItemsFromAllDrives: true})

  const templateData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Templates').getDataRange().getValues()

  const districtData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Districts').getDataRange().getValues().filter(r => r[0] !== '')
  const groupsData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Groups').getDataRange().getValues().filter(r => r[0] !== '')

  const year = new Date().getFullYear()

  //no headers
  districtData.shift()
  groupsData.shift()
  templateData.shift()

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

    const permissions = Drive.Permissions.list(districtFolder.id, {supportsAllDrives: true, includeItemsFromAllDrives: true, fields:'*'})

    if (permissions.permissions.find(e => e.emailAddress === email)) {
      console.log('District folder already shared with', email)
    } else {
      Drive.Permissions.create({type: 'user', role: 'writer', emailAddress: email}, districtFolder.id, {supportsAllDrives: true, includeItemsFromAllDrives: true})
      console.log('Sharing District folder with', email)
    }

    console.log('District folder created, copying templates', districtFolder.id)

    templateData.filter(r => r[1] === "district" && r[2] === "copy").forEach(template => {
      const file = Drive.Files.get(getIdFromUrl(template[0]), {supportsAllDrives: true, includeItemsFromAllDrives: true})
      const newName = applyTemplate(template[3], [['%F', file.name], ['%D', name], ['%N', number], ['%Y', year]])

      createOrCopyFile(file, newName, districtFolder)
    })

    templateData.filter(r => r[1] === "district" && r[2] === "shortcut").forEach(template => {
      const file = Drive.Files.get(getIdFromUrl(template[0]), {supportsAllDrives: true, includeItemsFromAllDrives: true})
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

    const permissions = Drive.Permissions.list(groupFolder.id, {supportsAllDrives: true, includeItemsFromAllDrives: true, fields:'*'})

    if (permissions.permissions.find(e => e.emailAddress === email)) {
      console.log('Group folder already shared with', email)
    } else {
      Drive.Permissions.create({type: 'user', role: 'writer', emailAddress: email}, groupFolder.id, {supportsAllDrives: true, includeItemsFromAllDrives: true})
      console.log('Sharing Group folder with', email)
    }

    console.log('Group folder created, copying templates', groupFolder.id)

    const copiedFiles = []

    templateData.filter(r => r[1] === "group" && r[2] === "copy").forEach(template => {
      const file = Drive.Files.get(getIdFromUrl(template[0]), {supportsAllDrives: true, includeItemsFromAllDrives: true})
      const newName = applyTemplate(template[3], [['%F', file.getName()], ['%G', name], ['%N', number], ['%D', district], ['%Y', year]])

      createOrCopyFile(file, newName, groupFolder, copiedFiles, template[4])
    })

    templateData.filter(r => r[1] === "group" && r[2] === "shortcut").forEach(template => {
      const file = Drive.Files.get(getIdFromUrl(template[0]), {supportsAllDrives: true, includeItemsFromAllDrives: true})
      createShortcutIfMissing(file, groupFolder)
    })

    console.log('Templates copied')
  })
}

function createOrGetFolder(name, parent) {
  const query = `name = '${name}' and '${parent.getId()}' in parents and trashed = false`
  const list = Drive.Files.list({ q: query, supportsAllDrives: true, includeItemsFromAllDrives: true, fields: "files(id, name, mimeType)" })

  if (list.files.length > 0) {
    console.log('Folder already exists', name)
    return list.files[0]
  } else {
    console.log('Creating folder', name)
    return Drive.Files.create({name: name, mimeType: 'application/vnd.google-apps.folder', parents: [parent.id]}, null, {supportsAllDrives: true})
  }
}

function createOrCopyFile(file, newName, parent, copiedFiles = [], flags) {
  const query = `name = '${newName}' and '${parent.getId()}' in parents and trashed = false`
  const list = Drive.Files.list({ q: query, supportsAllDrives: true, includeItemsFromAllDrives: true, fields: "files(id, name, mimeType)" })

  if (list.files.length > 0) {
    console.log('File already exists', newName)
    copiedFiles.push(list.files[0])
    return list.files[0]
  } else {
    console.log('Creating file', newName)
    const newFile = Drive.Files.copy({name: newName, parents:[parent.id]}, file.id, {supportsAllDrives: true, includeItemsFromAllDrives: true})
    copiedFiles.push(newFile)
    if(flags & 1) {
      const form = FormApp.openById(newFile.id)
      const spreadsheet = copiedFiles[copiedFiles.length - 2]
      form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.id)
      const responsesSheet = SpreadsheetApp.openById(spreadsheet.id)
      const formResponses = responsesSheet.getSheetByName("Form Responses 1").activate()
      console.log('Current index of sheet: %s', formResponses.getIndex());
      responsesSheet.moveActiveSheet(2)
      console.log('New index of sheet: %s', formResponses.getIndex());
    }
    if(flags & 2) {
      const spreadsheetToLoadFrom = copiedFiles[copiedFiles.length - 3]
      const register = SpreadsheetApp.openById(newFile.id)
      register.getActiveSheet().getRange(1,2).setValues([[spreadsheetToLoadFrom.id]])
    }
  }
}

function createShortcutIfMissing(file, parent) {
  const query = `name = '${file.name}' and '${parent.getId()}' in parents and trashed = false`
  const list = Drive.Files.list({ q: query, supportsAllDrives: true, includeItemsFromAllDrives: true, fields: "files(id, name, mimeType)" })

  if (list.files.length > 0) {
    console.log('Shortcut already exists for ', file.getName())
  } else {
    console.log('Creating shortcut', file.getName())
    Drive.Files.create({'name': file.name,'mimeType': 'application/vnd.google-apps.shortcut', parents: [parent.id], 'shortcutDetails': {'targetId': file.id}}, null, {supportsAllDrives: true, includeItemsFromAllDrives: true})
  }
}

function applyTemplate(template, substitutions) {
  return substitutions.reduce((acc, sub) => acc.replace(sub[0], sub[1]), template)
}

function getConfigValue(configData, item) {
  return configData.find(i => i[0] === item)[1]
}

function recursiveDriveList(root) {
  const list = Drive.Files.list({ q: `'${root}' in parents and trashed = false`, supportsAllDrives: true, includeItemsFromAllDrives: true, fields: "files(id, name, mimeType, permissionIds)" })
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