// Original typescript can be found at: https://github.com/vilhelm-k/Renomate-Kostnadskalkyl
// Use npm package @types/google-apps-script to get types for Google Apps Script

/**
 * @OnlyCurrentDoc
 */

const CONFIG_SHEET = 'Konfigurera projekt'
const DASHBOARD_SHEET = 'Dashboard'

const ADD_ROOMS_RANGE = 'AddRooms'
const CONFIG_EXISTING_ROOMS_RANGE = 'ConfigExistingRooms'
const DASHBOARD_MATERIAL_ROW_RANGE = 'DashboardMaterialRow'
const DASHBOARD_SUM_ROW_RANGE = 'DashboardSumRow'

const RENOMATE_YELLOW = '#fcd241'
const BUDGETING_TYPES = {
  'Enkel (rekommenderad)': '(enkel)',
  'Avancerad': '(avancerad)'
}

//############################################################################################################
//############################################# ADD ROOMS ####################################################
//############################################################################################################

type NewRoomRow = [string, string, keyof typeof BUDGETING_TYPES | '']

/**
 * Creates a list of pairs of the name of the new room and the template to copy from
 * @param newRooms: [[roomName, template, type]]
 * @param sheetNames: names of all sheets in the spreadsheet
 * @returns [[roomName, template]]
 * @throws Error if roomName, template or type is empty or if roomName already exists or there are duplicate roomNames
 */
const createRoomPairs = (newRooms: NewRoomRow[], sheetNames: string[]) => {
  let error = ''
  const roomPairs: [string, string][] = newRooms.map(([roomName, template, type], index) => {
    if (roomName === '') error += `Namn saknas till rum ${index + 1}\n`
    if (template === '') error += `Mall saknas till rum ${index + 1}\n`
    if (type === '') error += `Budgeteringsalternativ saknas till rum ${index + 1}\n`
    if (sheetNames.includes(roomName)) error += `Rum ${roomName} finns redan\n`
    return [roomName, `${template} ${BUDGETING_TYPES[type as keyof typeof BUDGETING_TYPES]}`]
  })
  const duplicates = roomPairs
    .map(([roomName]) => roomName)
    .filter((roomName, index, arr) => arr.indexOf(roomName) !== index)
  if (duplicates.length > 0) error += `Rumnamn upprepas: ${duplicates.join(', ')}\n`
  if (error !== '') throw new Error(error)
  return roomPairs
}

/**
 * Creates new sheets for the new rooms
 * @param roomTemplatePairs: [roomName, template][]
 */
const createNewRoomSheets = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, roomTemplatePairs: [string, string][]) => {
  for (const [roomName, template] of roomTemplatePairs)  // create new sheets
    ss.getSheetByName(template)?.copyTo(ss).setName(roomName).setTabColor(RENOMATE_YELLOW).showSheet()
}

/**
 * Adds new rooms to the dashboard
 * @param roomNames: names of the new rooms
 */
const addNewRoomsToDashboard = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, roomNames: string[]) => {
const dashboardSheet = <GoogleAppsScript.Spreadsheet.Sheet> ss.getSheetByName(DASHBOARD_SHEET)
const dashboardSumRowRange = <GoogleAppsScript.Spreadsheet.Range> ss.getRangeByName(DASHBOARD_SUM_ROW_RANGE)
  const insertRow = dashboardSumRowRange.getRow()

  const richTextValues = roomNames.map(roomName =>  // each room gets a link to its sheet
    [SpreadsheetApp.newRichTextValue()
      .setText(roomName)
      .setLinkUrl('#gid=' + ss.getSheetByName(roomName)?.getSheetId()) // maybe quicker with `'${roomName}'!A1`?
      .setTextStyle(SpreadsheetApp.newTextStyle().setBold(false).build())
      .build()]
  )
  dashboardSheet.insertRows(insertRow, roomNames.length)
  dashboardSheet.getRange(insertRow, 1, roomNames.length, 1).setRichTextValues(richTextValues)
}

/**
 * Adds new rooms to the spreadsheet
 * Both needs to create the new sheets and add the new rooms to the dashboard
 * Also checks so that the new rooms exist and are not duplicates, and that all values are filled in
 */
const addNewRooms = () => {
  const ss = SpreadsheetApp.getActive()
  const ui = SpreadsheetApp.getUi()
  try {
    const addRoomsRange = <GoogleAppsScript.Spreadsheet.Range> ss.getRangeByName(ADD_ROOMS_RANGE)
    const newRooms = <NewRoomRow[]> addRoomsRange.getValues()
      .filter(roomRow => roomRow.join('') !== '') // [roomName, template, type]
    const sheetNames = ss.getSheets().map(sheet => sheet.getName()) 
    const roomTemplatePairs = createRoomPairs(newRooms, sheetNames)
    if (roomTemplatePairs.length === 0) {
      ss.toast('Det finns inga nya rum att lägga till')
      return
    }
    createNewRoomSheets(ss, roomTemplatePairs)
    addNewRoomsToDashboard(ss, newRooms.map(([e]) => e))
    addRoomsRange.clearContent()
  } catch (err) {
    if (err instanceof Error) ui.alert(err.message)
  }
}

//############################################################################################################
//############################################ RENAME ROOMS ##################################################
//############################################################################################################

/**
 * Asks the user for new names for the rooms
 * @param oldRoomNames: names of the rooms to be renamed
 * @param sheetNames: names of all sheets in the spreadsheet
 * @returns : { [oldName: string]: string } or null if user pressed cancel
 */
const requestNewRoomNames = (oldRoomNames: string[], sheetNames: string[]) => {
  const ui = SpreadsheetApp.getUi()
  let renameMap = new Map<string, string>()
  for (const oldName of oldRoomNames) {
    while (true) {
      const response = ui.prompt(`Vad vill du döpa om ${oldName} till?`, ui.ButtonSet.OK_CANCEL)
      if (response.getSelectedButton() !== ui.Button.OK) return null
      const newRoomName = response.getResponseText().trim()

      if (newRoomName === '') ui.alert(`Namn saknas till rum ${oldName}`)
      else if (sheetNames.includes(newRoomName)) ui.alert(`Rum ${newRoomName} finns redan`)
      else if ([...renameMap.values()].includes(newRoomName)) ui.alert(`Du har redan angett: ${newRoomName}`)
      else {
        renameMap.set(oldName, newRoomName)
        break
      }
    }
  }
  return renameMap
}


/**
 * Renames the sheets as a part of the renameRooms function
 */
const renameSheets = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, renameMap: Map<string, string>) => {
  for (const [oldName, newName] of renameMap) {
    ss.getSheetByName(oldName)?.setName(newName)
  }
}

/**
 * Renames the rooms in the dashboard as a part of the renameRooms function
 */
const renameInDashboard = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, renameMap: Map<string, string>) => {
  const dashboardSheet = <GoogleAppsScript.Spreadsheet.Sheet> ss.getSheetByName(DASHBOARD_SHEET)
  const dashboardMaterialRow = <number> ss.getRangeByName(DASHBOARD_MATERIAL_ROW_RANGE)?.getRow()
  const dashboardSumRow = <number> ss.getRangeByName(DASHBOARD_SUM_ROW_RANGE)?.getRow()
  const dashboardRoomsRange = dashboardSheet.getRange(dashboardMaterialRow + 1, 1, dashboardSumRow - dashboardMaterialRow - 1, 1)
  const newRichText = dashboardRoomsRange.getRichTextValues().map(([room]) => {
    if (room === null) return [SpreadsheetApp.newRichTextValue().build()] // will never occur (probably). Just to make typescript happy
    const roomText = room.getText()
    if (!renameMap.has(roomText)) return [room]
    return [SpreadsheetApp.newRichTextValue()
      .setText(renameMap.get(roomText) as string)
      .setLinkUrl(room.getLinkUrl())
      .build()]
  })
  dashboardRoomsRange.setRichTextValues(newRichText)
}

/**
 * Renames the rooms selected in the config sheet
 * Asks the user for new names for the rooms via prompt
 * Renames the sheets and the rooms in the dashboard
 */
const renameRooms = () => {
  const ss = SpreadsheetApp.getActive()
  const configExistingRoomsRange = <GoogleAppsScript.Spreadsheet.Range> ss.getRangeByName(CONFIG_EXISTING_ROOMS_RANGE)
  const selectedRooms = configExistingRoomsRange.getValues()
    .filter(([checkbox, ]) => checkbox)
    .map(([, roomName]) => roomName)
  if (selectedRooms.length === 0) {
    ss.toast('Inga rum valda')
    return
  }
  const sheetNames = ss.getSheets().map(sheet => sheet.getName())
  
  const renamedRooms = requestNewRoomNames(selectedRooms, sheetNames)
  if (renamedRooms === null) return // user pressed cancel

  renameSheets(ss, renamedRooms)
  renameInDashboard(ss, renamedRooms)
  configExistingRoomsRange.uncheck()
}

//############################################################################################################
//############################################# DELETE ROOMS #################################################
//############################################################################################################

/**
 * Deletes the sheets for the given rooms
 * @param roomNames: names of the rooms to be deleted
 */
const deleteSheets = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, roomNames: string[]) => {
  for (const roomName of roomNames) {
    const deleteSheet = ss.getSheetByName(roomName)
    if (deleteSheet === null) continue
    ss.deleteSheet(deleteSheet)
  }
}

/**
 * Groups adjacent numbers into ranges
 * Also reverses the order of the groups so that they are correctly deleted
 * @param numbers: numbers to be grouped. Assumed to be sorted
 */
const groupAdjacentNumbers = (numbers: number[]) => 
  numbers.reduce((groups, num, i, arr) => {
    if (i === 0 || num !== arr[i - 1] + 1) {
      groups.push([num, 1])
    } else groups[groups.length - 1][1] += 1
    return groups
    }, [] as [number, number][]).reverse()


/**
 * Deletes the rooms in the dashboard
 * @param roomNames: names of the rooms to be deleted
 */
const deleteInDashboard = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, roomNames: string[]) => {
  const dashboardSheet = <GoogleAppsScript.Spreadsheet.Sheet> ss.getSheetByName(DASHBOARD_SHEET)
  const firstColumnValues = dashboardSheet.getRange(1, 1, dashboardSheet.getLastRow(), 1).getValues()
  const deleteIndexes = roomNames
    .map(roomName => firstColumnValues.findIndex(([row]) => row === roomName) + 1)
    .filter(index => index !== 0) // since we added 1 to the index
  const deleteGroups = groupAdjacentNumbers(deleteIndexes)
  for (const [position, numRows] of deleteGroups) dashboardSheet.deleteRows(position, numRows)
}
    
/**
 * Deletes the rooms selected in the config sheet
 */
const deleteRooms = () => {
  const ss = SpreadsheetApp.getActive()
  const ui = SpreadsheetApp.getUi()
  const configExistingRoomsRange = <GoogleAppsScript.Spreadsheet.Range> ss.getRangeByName(CONFIG_EXISTING_ROOMS_RANGE)
  const selectedRooms = configExistingRoomsRange.getValues()
    .filter(([checkbox, ]) => checkbox)
    .map(([, roomName]) => roomName)
  if (selectedRooms.length === 0) {
    ss.toast('Inga rum valda')
    return
  }
  
  const response = ui.alert(`Vill du verkligen ta bort rummen: ${selectedRooms.join(', ')}?`, ui.ButtonSet.YES_NO)
  if (response !== ui.Button.YES) return

  deleteSheets(ss, selectedRooms)
  deleteInDashboard(ss, selectedRooms)
  configExistingRoomsRange.uncheck()
<<<<<<< HEAD
}
