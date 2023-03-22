// Original typescript can be found at: https://github.com/vilhelm-k/Renomate-Kostnadskalkyl
// Use npm package @types/google-apps-script to get types for Google Apps Script

/**
 * @OnlyCurrentDoc
 */

const DASHBOARD_SHEET = 'Sammanfattning';

const ADD_ROOMS_RANGE = 'AddRooms';
const CONFIG_EXISTING_ROOMS_RANGE = 'ConfigExistingRooms';
const DASHBOARD_MATERIAL_ROW_RANGE = 'DashboardMaterialRow';
const DASHBOARD_SUM_ROW_RANGE = 'DashboardSumRow';

const RENOMATE_YELLOW = '#fcd241';

type StringPair = [string, string];

// ############################################################################################################
// ########################################### ON STARTUP #####################################################
// ############################################################################################################

/** Used for user journey */
const activateScripts = () => SpreadsheetApp.getActive().toast('Skriptet är redan aktiverat');

// ############################################################################################################
// ############################################# ADD ROOMS ####################################################
// ############################################################################################################

/**
 * Creates a list of pairs of the name of the new room and the template to copy from
 * @param newRoom [roomName, template, type]
 * @param sheetNames names of all sheets in the spreadsheet
 * @returns [roomName, template]
 * @throws Error if roomName, template or type is empty or if roomName already exists
 */
const createRoomNameAndTemplatePair = ([inputRoomName, template]: StringPair, sheetNames: string[]): StringPair => {
  const errors: string[] = [];
  const roomName = inputRoomName.trim();

  if (roomName === '') errors.push('Namn saknas');
  if (template === '') errors.push('Mall saknas');
  if (sheetNames.includes(roomName)) errors.push(`Rum ${roomName} finns redan`);
  if (errors.length > 0) throw new Error(errors.join('\n'));

  return [roomName, `${template} (mall)`];
};

/**
 * Creates new sheets for the new rooms
 * @param nameTemplatePairs [roomName, template][]
 */
const createNewRoomSheet = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, [roomName, template]: StringPair) =>
  ss.getSheetByName(template)?.copyTo(ss).setName(roomName).setTabColor(RENOMATE_YELLOW).showSheet();

/**
 * Adds new rooms to the dashboard
 * @param roomNames names of the new rooms
 */
const addNewRoomNamesToDashboard = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, roomName: string) => {
  const dashboardSheet = <GoogleAppsScript.Spreadsheet.Sheet>ss.getSheetByName(DASHBOARD_SHEET);
  const dashboardSumRowRange = <GoogleAppsScript.Spreadsheet.Range>ss.getRangeByName(DASHBOARD_SUM_ROW_RANGE);
  const insertRow = dashboardSumRowRange.getRow();

  const richTextValue = SpreadsheetApp.newRichTextValue()
    .setText(roomName)
    .setLinkUrl(`#gid=${ss.getSheetByName(roomName)?.getSheetId()}`) // maybe quicker with `'${roomName}'!A1`?
    .setTextStyle(SpreadsheetApp.newTextStyle().setBold(false).build())
    .build();

  dashboardSheet.insertRows(insertRow, 1);
  dashboardSheet.getRange(insertRow, 1, 1, 1).setRichTextValue(richTextValue);
};

/**
 * Adds new rooms to the spreadsheet
 * Both needs to create the new sheets and add the new rooms to the dashboard
 * Also checks so that the new rooms exist and are not duplicates, and that all values are filled in
 */
const addNewRooms = () => {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  try {
    const addRoomsRange = <GoogleAppsScript.Spreadsheet.Range>ss.getRangeByName(ADD_ROOMS_RANGE);
    const newRoomRow = <StringPair>addRoomsRange.getValues().filter((roomRow) => roomRow.join('') !== '')[0];
    if (newRoomRow[0] === '' && newRoomRow[1] === '') {
      ss.toast('Inga nya rum att lägga till');
      return;
    }
    const sheetNames = ss.getSheets().map((sheet) => sheet.getName());
    const nameTemplatePair = createRoomNameAndTemplatePair(newRoomRow, sheetNames);

    createNewRoomSheet(ss, nameTemplatePair);
    addNewRoomNamesToDashboard(ss, nameTemplatePair[0]);
    addRoomsRange.clearContent();
  } catch (err) {
    if (err instanceof Error) ui.alert(err.message);
  }
};

// ############################################################################################################
// ############################################ RENAME ROOMS ##################################################
// ############################################################################################################

/**
 * Gets rooms selected in config sheet
 */
const getSelectedRoomNames = (configExistingRoomsRange: GoogleAppsScript.Spreadsheet.Range) =>
  configExistingRoomsRange
    .getValues()
    .filter(([, , checkbox]) => checkbox)
    .map(([roomName]) => roomName);

/**
 * Asks the user for new names for the rooms. Ensures that names are valid
 * @param oldRoomNames names of the rooms to be renamed
 * @param sheetNames names of all sheets in the spreadsheet
 * @returns Map<oldName, newName> or null if user pressed cancel
 */
const requestNewRoomNames = (oldRoomNames: string[], sheetNames: string[]) => {
  const ui = SpreadsheetApp.getUi();
  const renameMap = new Map<string, string>();
  for (const oldName of oldRoomNames) {
    while (true) {
      const response = ui.prompt(`Vad vill du döpa om ${oldName} till?`, ui.ButtonSet.OK_CANCEL);
      if (response.getSelectedButton() !== ui.Button.OK) return null;
      const newRoomName = response.getResponseText().trim();

      if (newRoomName === '') ui.alert(`Namn saknas till rum ${oldName}`);
      else if (sheetNames.includes(newRoomName)) ui.alert(`Rum ${newRoomName} finns redan`);
      else if ([...renameMap.values()].includes(newRoomName)) ui.alert(`Du har redan angett: ${newRoomName}`);
      else {
        renameMap.set(oldName, newRoomName);
        break;
      }
    }
  }
  return renameMap;
};

/**
 * Renames the sheets as a part of the renameRooms function
 */
const renameSheets = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, renameMap: Map<string, string>) => {
  for (const [oldName, newName] of renameMap) ss.getSheetByName(oldName)?.setName(newName);
};

/**
 * Renames the rooms in the dashboard as a part of the renameRooms function
 */
const renameInDashboard = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, renameMap: Map<string, string>) => {
  const dashboardSheet = <GoogleAppsScript.Spreadsheet.Sheet>ss.getSheetByName(DASHBOARD_SHEET);
  const materialRow = <number>ss.getRangeByName(DASHBOARD_MATERIAL_ROW_RANGE)?.getRow();
  const sumRow = <number>ss.getRangeByName(DASHBOARD_SUM_ROW_RANGE)?.getRow();
  const roomsRange = dashboardSheet.getRange(materialRow + 1, 1, sumRow - materialRow - 1, 1);
  const newRichText = roomsRange.getRichTextValues().map(([room]) => {
    if (room === null) return [SpreadsheetApp.newRichTextValue().build()]; // will never occur (probably)
    const roomText = room.getText();
    if (!renameMap.has(roomText)) return [room];
    return [
      SpreadsheetApp.newRichTextValue()
        .setText(renameMap.get(roomText) as string)
        .setLinkUrl(room.getLinkUrl())
        .build(),
    ];
  });
  roomsRange.setRichTextValues(newRichText);
};

/**
 * Renames the rooms selected in the config sheet
 * Asks the user for new names for the rooms via prompt
 * Renames the sheets and the rooms in the dashboard
 */
const renameRooms = () => {
  const ss = SpreadsheetApp.getActive();
  const configExistingRoomsRange = <GoogleAppsScript.Spreadsheet.Range>ss.getRangeByName(CONFIG_EXISTING_ROOMS_RANGE);
  const selectedRooms = getSelectedRoomNames(configExistingRoomsRange);

  if (selectedRooms.length === 0) {
    ss.toast('Inga rum valda');
    return;
  }
  const sheetNames = ss.getSheets().map((sheet) => sheet.getName());

  const renamedRooms = requestNewRoomNames(selectedRooms, sheetNames);
  if (renamedRooms === null) return; // user pressed cancel

  renameSheets(ss, renamedRooms);
  renameInDashboard(ss, renamedRooms);
  configExistingRoomsRange.uncheck();
};

// ############################################################################################################
// ############################################# DELETE ROOMS #################################################
// ############################################################################################################

/**
 * Deletes the sheets for the given rooms
 * @param roomNames names of the rooms to be deleted
 */
const deleteSheets = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, roomNames: string[]) => {
  for (const roomName of roomNames) {
    const deleteSheet = ss.getSheetByName(roomName);
    if (deleteSheet !== null) ss.deleteSheet(deleteSheet);
  }
};

/**
 * Groups adjacent numbers into ranges for more efficient deletion in the dashboard
 * Also reverses the order of the groups so that they are correctly deleted
 * @param numbers numbers to be grouped. Has to be sorted in ascending order
 */
const groupAdjacentNumbers = (numbers: number[]) => {
  const groups: [number, number][] = [];
  for (let i = 0; i < numbers.length; i++) {
    if (i === 0 || numbers[i] !== numbers[i - 1] + 1) groups.push([numbers[i], 1]);
    else groups[groups.length - 1][1] += 1;
  }
  return groups.reverse();
};

/**
 * Deletes the rooms in the dashboard
 * @param roomNames names of the rooms to be deleted
 */
const deleteInDashboard = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, roomNames: string[]) => {
  const dashboardSheet = <GoogleAppsScript.Spreadsheet.Sheet>ss.getSheetByName(DASHBOARD_SHEET);
  const firstColumnValues = dashboardSheet.getRange(1, 1, dashboardSheet.getLastRow(), 1).getValues();
  const deleteIndexes = roomNames
    .map((roomName) => firstColumnValues.findIndex(([row]) => row === roomName) + 1)
    .filter((index) => index !== 0); // since we added 1 to the index
  const deleteGroups = groupAdjacentNumbers(deleteIndexes);
  for (const [position, numRows] of deleteGroups) dashboardSheet.deleteRows(position, numRows);
};

/**
 * Deletes the rooms selected in the config sheet
 * Deletes both their corresponding sheets and their row in the dashboard
 */
const deleteRooms = () => {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const configExistingRoomsRange = <GoogleAppsScript.Spreadsheet.Range>ss.getRangeByName(CONFIG_EXISTING_ROOMS_RANGE);
  const selectedRooms = getSelectedRoomNames(configExistingRoomsRange);
  if (selectedRooms.length === 0) {
    ss.toast('Inga rum valda');
    return;
  }

  const response = ui.alert(`Vill du verkligen ta bort rummen: ${selectedRooms.join(', ')}?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  deleteSheets(ss, selectedRooms);
  deleteInDashboard(ss, selectedRooms);
  configExistingRoomsRange.uncheck();
};
