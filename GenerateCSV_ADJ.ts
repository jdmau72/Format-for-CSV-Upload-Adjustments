// author: Justin D. Mau
// script parses the document to create a list of adjustment objects
// it then generates a csv that can be uploaded to Netsuite to process all adjustments quickly


class Adjustment {
  referenceNumber: string;
  adjustmentAccount: string;
  adjustmentAccountReason: string;
  binNumber: string;
  item: string;
  location: string;
  adjQtyBy: number;
  lotNumber: string;
  quantity: number;
  division: string;
  department: string;
  date: string;
  memo: string;

  category: string;

  constructor(item: string, lot: string, quantity: number, bin: string, location: string, memo: string){
    this.referenceNumber = "To Be Generated";

    // now check if the item is an implant or instrument
    this.category = findCategory(item);
    if (this.category == 'instrument') {
      this.adjustmentAccount = "65390 Manufacturing Overhead : Instrument wear";
      this.adjustmentAccountReason = "SI Transaction";
    } else if (this.category == 'implant') {
      this.adjustmentAccount = "65240 Manufacturing Overhead : Physical inventory adjustments";
      this.adjustmentAccountReason = "HW Transaction";
    }

    this.binNumber = bin;
    this.item = item;
    this.location = location;
    this.adjQtyBy = this.quantity = quantity;
    this.lotNumber = lot;
    this.division = "Hardware";
    this.department = "265-Other OH - HW";
    this.date = new Date().toLocaleDateString();
    this.memo = memo;
  }

    getValues(): string[]{
      return [this.referenceNumber, this.adjustmentAccount, this.adjustmentAccountReason,
                this.binNumber, this.item, this.location, this.adjQtyBy.toString(), this.lotNumber, this.quantity.toString(), this.division, this.department, this.date, this.memo];
    }

    getCategory(): string {
      return this.category;
    }

}


// creates a list of adjustments accessible for all functions in the code
let adjustmentList: Adjustment[] = [];



// MAIN -------------------------------------------------------------------------------------------------------------
function main(workbook: ExcelScript.Workbook,
  location: string = "TRAYBUILD",
  memo: string = "Cortera MMO Trays - ") {

  // Get the active cell and worksheet.
  let selectedCell = workbook.getActiveCell();
  let sheet = workbook.getFirstWorksheet();

  let usedRange = sheet.getUsedRange();
  let numRows = usedRange.getValues().length;


  // first gets the column indices for each required value
  let headerRange = sheet.getRange("A1:Z1");
  let adjColIndex = findColumn("LOT ADJ", headerRange);
  let lotColIndex = findColumn("Lot", headerRange);
  let itemColIndex = findColumn("Item Number", headerRange);
  let qtyColIndex = findColumn("QTY", headerRange);
  let binColIndex = findColumn("Bin Number", headerRange);

  // gets the bin number
  let binNumber = sheet.getCell(1, binColIndex).getValue() as string;

  // now loop through each item
  for (let row = 1; row < numRows; row++){

    // gets the values of each important category
    let rowValues = sheet.getRangeByIndexes(row, 0, 1, 99).getValues()[0];
    let adj = rowValues[adjColIndex] as string;
    let lot = rowValues[lotColIndex] as string;
    let item = rowValues[itemColIndex] as string;
    let qty = rowValues[qtyColIndex] as number;

  // if the adj col is empty, ignore
    if (adj != ""){
      
      // checks to see if there are multiple lots for adjustments
      let subAdjustments = adj.toString().split(","); // tries to split if there are multiple lots for that one item
      if (subAdjustments.length > 1) {

        for (let i = 0; i < subAdjustments.length; i++)
        {
          let subAdj = subAdjustments[i].trim().split(" "); // splits each of those lots to get the qty and lot
          let subAdjLot = subAdj[0].trim();
          let subAdjQty = Number(subAdj[1].replace(/[()]/g, ''));

          // creates an adjustment entry and adds to the list
          let adjOut = new Adjustment(item, lot, (subAdjQty * -1), binNumber, location, memo);  // one to remove the original lot number
          let adjIn = new Adjustment(item, subAdjLot, subAdjQty, binNumber, location, memo);  // one to add the new one
          adjustmentList.push(adjOut);
          adjustmentList.push(adjIn);
        }

      // otherwise it will just do the one lot adjustment
      } else
       {
        // will try to split in case only some need the lot adjustment
        let subAdj = adj.toString().trim().split(" "); // splits each of those lots to get the qty and lot

        if (subAdj.length > 1) // if there is a quantity listed in the adj column
        {
          adj = subAdj[0].trim(); // replaces adjustment lot with the split lot
          qty = Number(subAdj[1].replace(/[()]/g, '')); // replaces qty with the specified amount
        }

        let adjOut = new Adjustment(item, lot, (qty * -1), binNumber, location, memo);
        let adjIn = new Adjustment(item, adj, qty, binNumber, location, memo);
        adjustmentList.push(adjOut);
        adjustmentList.push(adjIn);
      }
    }
  }

  console.log(adjustmentList);

  // now generates a new sheet for those adjustments and formats it to fit the CSV upload
  generateAdjustmentCSV(workbook, binNumber, "implant");
  generateAdjustmentCSV(workbook, binNumber, "instrument");
}
// end of MAIN ---------------------------------------------------------------------------------------------------------



// HELPER FUNCTIONS ----------------------------------------------------------------------------------------
function findColumn(searchTerm: string, headerRange: ExcelScript.Range){
    // gets the values of the headers
    let headers = headerRange.getValues()[0];
    let adjCol = 0;

    for (let col = 0; col < headers.length; col++){
        if (headers[col].toString().toLowerCase() == searchTerm.toLowerCase()) {
          adjCol = col;
        }
    }
    return adjCol;
}


function generateAdjustmentCSV(workbook: ExcelScript.Workbook, bin: string, category: string){
  
  // create new sheet for the adjustments of that category
  let adjSheet = workbook.addWorksheet(`${bin}-ADJ-${category.toUpperCase()}`);

  // defines the headers
  let headerRange = adjSheet.getRangeByIndexes(0, 0, 1, 13);
  headerRange.setValues([["Reference #", "Adjustment Account", "Adjustment Account Reason", 	"Bin Number",	"Item",	"Location",	"Adjust Qty. By",	"Receipt Inventory Number",	"Quantitiy", 	"Division",	"Department", "Date", "Memo"]])

  // loop through adj list, if category matches, add it to the sheet
  let row = 1;
  for (let i = 1; i < adjustmentList.length + 1; i++){
    if (adjustmentList[i - 1].getCategory().toUpperCase() == category.toUpperCase()){
      let adj = adjustmentList[i - 1].getValues();
      adjSheet.getRangeByIndexes(row, 0, 1, 13).setValues([adj]);
      row++;
    }
  }
}



// findCategory is currently very basic and is just for Cortera products
// in the future, adding a database of instrument and implant codes would help extend to other products
function findCategory(item: string){
  if (item.substring(0, 4) == "1509" || item.substring(0, 4) == "1505"){
    return "instrument";
  } else {
    return "implant";
  }
}
// end of HELPER FUNCTIONS -------------------------------------------------------------------------------
