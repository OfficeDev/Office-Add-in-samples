import {
  SetStartupBehaviorHelper,
  SetRuntimeVisibleHelper,
  updateRibbon,
  getGlobal,
  connectService,
  monitorSheetChanges
} from '../../utilities/office-apis-helpers';

const g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.btnenableaddinstart = btnEnableAddinStart;
g.btndisableaddinstart = btnDisableAddinStart;
g.btninsertdata = btnInsertData;
g.btnopentaskpane = btnOpenTaskpane;
g.btnclosetaskpane = btnCloseTaskpane;
g.btnconnectservice = btnConnectService;
g.btndisconnectservice = btnDisconnectService;
g.btnsumdata = btnSumData;

export function btnConnectService(event: Office.AddinCommands.Event) {
  console.log('Connect service button pressed');
  // Your code goes here
  g.state.setConnected(true);
  g.state.isConnectInProgress = true;
  updateRibbon();
  connectService();
  monitorSheetChanges();
  event.completed();
}
export function btnDisconnectService(event: Office.AddinCommands.Event) {
  console.log('Disconnect service button pressed');
  // Your code goes here
  g.state.setConnected(false);
  updateRibbon();
  event.completed();
}

export function btnOpenTaskpane(event: Office.AddinCommands.Event) {
  console.log('Open task pane button pressed');
  // Your code goes here
  SetRuntimeVisibleHelper(true);
  g.state.isTaskpaneOpen = true;
  updateRibbon();
  event.completed();
}

export function btnCloseTaskpane(event: Office.AddinCommands.Event) {
  console.log('Open task pane button pressed');
  // Your code goes here
  SetRuntimeVisibleHelper(false);
  g.state.isTaskpaneOpen = false;
  updateRibbon();
  event.completed();
}

export function btnEnableAddinStart(event: Office.AddinCommands.Event) {
  console.log('Enable add-in start button pressed');
  // Your code goes here
  SetStartupBehaviorHelper(true);
  g.state.isStartOnDocOpen = true;
  updateRibbon();
  event.completed();
}

export function btnDisableAddinStart(event: Office.AddinCommands.Event) {
  console.log('Disable add-in start button pressed');
  // Your code goes here
  SetStartupBehaviorHelper(false);
  g.state.isStartOnDocOpen = false;
  updateRibbon();

  event.completed();
}

export function btnInsertData(event: Office.AddinCommands.Event) {
  console.log('Insert data button pressed');
  // Mock code that pretends to insert data from a data source
  insertData();
  event.completed();
}

export async function btnSumData(event: Office.AddinCommands.Event) {
  console.log('Insert data button pressed');
  // Mock code that pretends to insert data from a data source
  let address = g.state.selectionAddress as string;
  await Excel.run(context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange(address);
    range.load('values');

    let sum: number = 0;
    return context.sync().then(() => {
      range.values.forEach(v => {
        let vnumber: number = +v.toString();
        sum += vnumber;
      });

      return context.sync().then(() => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();

        let range = sheet.getRange('F1');
        range.values = [[sum]];
        range.format.autofitColumns();
        event.completed();
        console.log(sum);
        return context.sync();
      });
    });
  });
  event.completed();
}

async function insertData() {
  try {
    await Excel.run(async context => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.add('A1:D1', true /*hasHeaders*/);
      expensesTable.name = 'ExpensesTable';

      expensesTable.getHeaderRowRange().values = [
        ['Date', 'Merchant', 'Category', 'Amount']
      ];

      expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ['1/1/2017', 'The Phone Company', 'Communications', '$120'],
        ['1/2/2017', 'Northwind Electric Cars', 'Transportation', '$142'],
        ['1/5/2017', 'Best For You Organics Company', 'Groceries', '$27'],
        ['1/10/2017', 'Coho Vineyard', 'Restaurant', '$33'],
        ['1/11/2017', 'Bellows College', 'Education', '$350'],
        ['1/15/2017', 'Trey Research', 'Other', '$135'],
        ['1/15/2017', 'Best For You Organics Company', 'Groceries', '$97']
      ]);

      if (Office.context.requirements.isSetSupported('ExcelApi', '1.2')) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }
      context.sync().then(() => {
        monitorSheetChanges();
        return context.sync();
      });
    });
  } catch (error) {
    console.log(error);
  }
}
