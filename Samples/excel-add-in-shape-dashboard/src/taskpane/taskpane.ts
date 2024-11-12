/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("add-sample-data").onclick = addSampleData;
    document.getElementById("create-dash-board").onclick = createDashboard;
    document.getElementById("add-information").onclick = changeColor;
    document.getElementById("change-font-format").onclick = changeFontFormat;
  }
});

async function createDashboard() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      let shapes = context.workbook.worksheets.getItem("Sample").shapes;
      let array = [];

      let region = shapes.addGeometricShape("RoundRectangle");
      region.left = 550;
      region.top = 100;
      region.width = 100;
      region.height = 100;
      array.push(region);

      let state = shapes.addGeometricShape("RoundRectangle");
      state.left = 700;
      state.top = 100;
      state.width = 100;
      state.height = 100;
      array.push(state);

      let line = shapes.addGeometricShape("Rectangle");
      line.left = 500;
      line.top = 230;
      line.width = 350;
      line.height = 1;
      array.push(line);

      let category = shapes.addGeometricShape("RoundRectangle");
      category.left = 550;
      category.top = 260;
      category.width = 100;
      category.height = 100;
      array.push(category);

      let subCategory = shapes.addGeometricShape("RoundRectangle");
      subCategory.left = 700;
      subCategory.top = 260;
      subCategory.width = 100;
      subCategory.height = 100;
      array.push(subCategory);

      let line2 = shapes.addGeometricShape("Rectangle");
      line2.left = 500;
      line2.top = 390;
      line2.width = 350;
      line2.height = 1;
      array.push(line2);

      let maxScale = shapes.addGeometricShape("RoundRectangle");
      maxScale.left = 550;
      maxScale.top = 420;
      maxScale.width = 100;
      maxScale.height = 100;
      array.push(maxScale);

      let sumScale = shapes.addGeometricShape("RoundRectangle");
      sumScale.left = 700;
      sumScale.top = 420;
      sumScale.width = 100;
      sumScale.height = 100;
      array.push(sumScale);

      shapes.addGroup(array);
      await context.sync();
      
      showStatus('Create empty dashboard - success!', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function onDeactivate() {
  await changeColor();
  await changeFontFormat();
}

async function changeColor() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      let shapes = context.workbook.worksheets.getItem("Sample").shapes;
      shapes.load("items");
      await context.sync();

      let shapeGroup = shapes.items[0].group.shapes;
      shapeGroup.load("items");
      await context.sync();

      let text2 = context.workbook.worksheets.getItem("Sample").getRange("E152:E152");
      text2.load("values");
      await context.sync();

      let region = shapeGroup.items[0];
      region.textFrame.textRange.text = Math.ceil(parseFloat(text2.values.toString())).toString() + "\nRegions";
      region.fill.foreColor = "#339933";

      let text3 = context.workbook.worksheets.getItem("Sample").getRange("E157:E157");
      text3.load("values");
      await context.sync();
      let state = shapeGroup.items[1];
      state.fill.foreColor = "#339933";
      state.textFrame.textRange.text = text3.values.toString() + "\nStates";

      let text4 = context.workbook.worksheets.getItem("Sample").getRange("E153:E153");
      text4.load("values");
      await context.sync();
      let category = shapeGroup.items[3];
      category.fill.foreColor = "#003366";
      category.textFrame.textRange.text = text4.values.toString() + "\nCategories";

      let text5 = context.workbook.worksheets.getItem("Sample").getRange("E154:E154");
      text5.load("values");
      await context.sync();
      let subCategory = shapeGroup.items[4];
      subCategory.fill.foreColor = "#003366";
      subCategory.textFrame.textRange.text = text5.values.toString() + "\nSub Categories";

      let text6 = context.workbook.worksheets.getItem("Sample").getRange("E155:E155");
      text6.load("values");
      await context.sync();
      let maxSale = shapeGroup.items[6];
      maxSale.fill.foreColor = "#FF6600";
      maxSale.textFrame.textRange.text = text6.values.toString() + "\nMax Sale";

      let text7 = context.workbook.worksheets.getItem("Sample").getRange("E156:E156");
      text7.load("values");
      await context.sync();
      let sumSale = shapeGroup.items[7];
      sumSale.fill.foreColor = "#FF6600";
      sumSale.textFrame.textRange.text = text7.values.toString() + "\nSum Sale";

      await context.sync();
      showStatus('Add information to dashboard - success!', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function changeFontFormat() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      let shapes = context.workbook.worksheets.getItem("Sample").shapes;
      shapes.load("items");
      await context.sync();

      let shapeGroup = shapes.items[0].group.shapes;
      shapes.items[0].onActivated.add(onActivate);
      shapes.items[0].onDeactivated.add(onDeactivate);
      shapeGroup.load("items");
      await context.sync();

      for (let i = 0; i < shapeGroup.items.length; i++) {
        let shp = shapeGroup.items[i];
        shp.textFrame.textRange.font.name = "Consolas";

        shp.textFrame.verticalAlignment = "Middle";
        shp.textFrame.horizontalAlignment = "Center";
      }

      let region = shapeGroup.items[0];
      region.textFrame.textRange.getSubstring(0, 1).font.size = 30;
      region.textFrame.textRange.getSubstring(1).font.size = 17;
      region.textFrame.textRange.getSubstring(1).font.color = "#FFFFCC";

      let state = shapeGroup.items[1];
      state.textFrame.textRange.getSubstring(0, 2).font.size = 30;
      state.textFrame.textRange.getSubstring(2).font.size = 17;
      state.textFrame.textRange.getSubstring(2).font.color = "#FFFFCC";

      let category = shapeGroup.items[3];
      category.textFrame.textRange.getSubstring(0, 1).font.size = 30;
      category.textFrame.textRange.getSubstring(1).font.size = 12;
      category.textFrame.textRange.getSubstring(1).font.color = "#99CCFF";

      let subCategory = shapeGroup.items[4];
      subCategory.textFrame.textRange.getSubstring(0, 2).font.size = 30;
      subCategory.textFrame.textRange.getSubstring(2).font.size = 12;
      subCategory.textFrame.textRange.getSubstring(2).font.color = "#99CCFF";

      let maxSale = shapeGroup.items[6];
      maxSale.textFrame.textRange.getSubstring(0, 8).font.size = 18;
      maxSale.textFrame.textRange.getSubstring(8).font.size = 13;
      maxSale.textFrame.textRange.getSubstring(8).font.color = "FFFF66";

      let sumSale = shapeGroup.items[7];
      sumSale.textFrame.textRange.getSubstring(0, 9).font.size = 15;
      sumSale.textFrame.textRange.getSubstring(9).font.size = 13;
      sumSale.textFrame.textRange.getSubstring(9).font.color = "FFFF66";

      await context.sync();
      showStatus('Change information format - success!', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function onActivate() {
  try {
    await Excel.run(async (context) => {
      let shapes = context.workbook.worksheets.getItem("Sample").shapes;
      shapes.load("items");
      await context.sync();
      let shapeGroup = shapes.items[0].group.shapes;
      shapeGroup.load("items");
      await context.sync();

      let region = shapeGroup.items[0];
      region.textFrame.textRange.text = "-Central\n-South\n-West\n-East";
      region.textFrame.horizontalAlignment = "Left";
      region.textFrame.textRange.font.size = 13;

      let state = shapeGroup.items[1];
      state.textFrame.textRange.text = "Top 3:\n-California\n-Washington\n-New York";
      state.textFrame.horizontalAlignment = "Left";
      state.textFrame.textRange.font.size = 12;

      let category = shapeGroup.items[3];
      category.textFrame.textRange.text = "-OfficeSupply\n-Furniture\n-Technology";
      category.textFrame.horizontalAlignment = "Left";
      category.textFrame.textRange.font.size = 10;

      let subCategory = shapeGroup.items[4];
      subCategory.textFrame.textRange.text = "Top 3:\n-Phone\n-Tables\n-Accessories";
      subCategory.textFrame.horizontalAlignment = "Left";
      subCategory.textFrame.textRange.font.size = 10;

      let maxScale = shapeGroup.items[6];
      maxScale.textFrame.textRange.text = "US2015126214\n-Furniture\n-Washington";
      maxScale.textFrame.textRange.font.size = 11;

      let subScale = shapeGroup.items[7];
      subScale.textFrame.textRange.text = "Top 3:\nUS2015126214\nCA2017137596\nCA2014115812";
      subScale.textFrame.horizontalAlignment = "Left";
      subScale.textFrame.textRange.font.size = 11;

      await context.sync();
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function addSampleData() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      context.workbook.worksheets.getItemOrNullObject("Sample").delete();
      context.workbook.worksheets.add("Sample");
      const sheet = context.workbook.worksheets.getItem("Sample");
      sheet.activate();

      let expensesTable = sheet.tables.add("A1:H1", true);
      expensesTable.name = "SalesTable";
      expensesTable.getHeaderRowRange().values = [
        ["Order ID", "Ship Date", "Segment", "Region", "Category", "Sub-Category", "Sales", "State"],
      ];
      expensesTable.rows.add(null, [
        ["CA-2014-112326", 41647, "Home Office", "Central", "Office Supplies", "Labels", 11.784, "Illinois"],
        ["CA-2014-112326", 41647, "Home Office", "Central", "Office Supplies", "Storage", 272.736, "Illinois"],
        ["CA-2014-135405", 41652, "Consumer", "Central", "Office Supplies", "Art", 9.344, "Texas"],
        ["CA-2014-149020", 41654, "Corporate", "South", "Office Supplies", "Labels", 2.89, "Virginia"],
        ["CA-2014-155852", 41705, "Consumer", "South", "Office Supplies", "Art", 19.456, "North Carolina"],
        ["CA-2014-104472", 41797, "Home Office", "West", "Furniture", "Furnishings", 73.32, "Utah"],
        ["CA-2014-115812", 41804, "Consumer", "West", "Office Supplies", "Art", 7.28, "California"],
        ["CA-2014-115812", 41804, "Consumer", "West", "Technology", "Phones", 907.152, "California"],
        ["US-2014-141215", 41811, "Corporate", "Central", "Furniture", "Tables", 99.918, "Texas"],
        ["CA-2014-157784", 41828, "Consumer", "South", "Office Supplies", "Paper", 19.44, "Mississippi"],
        ["US-2014-119137", 41847, "Consumer", "West", "Office Supplies", "Art", 9.24, "Arizona"],
        ["CA-2014-113362", 41901, "Consumer", "East", "Office Supplies", "Storage", 449.15, "New York"],
        ["CA-2014-156601", 41906, "Corporate", "West", "Office Supplies", "Fasteners", 7.16, "California"],
        ["US-2014-134614", 41907, "Consumer", "Central", "Furniture", "Tables", 617.7, "Illinois"],
        ["CA-2014-134677", 41922, "Consumer", "West", "Technology", "Accessories", 9.09, "California"],
        ["CA-2014-139451", 41928, "Consumer", "West", "Office Supplies", "Art", 14.9, "California"],
        ["CA-2014-164973", 41952, "Home Office", "East", "Technology", "Accessories", 360, "New York"],
        ["CA-2014-144666", 41954, "Consumer", "West", "Furniture", "Bookcases", 22.666, "California"],
        ["CA-2014-163419", 41957, "Consumer", "West", "Technology", "Phones", 559.984, "Colorado"],
        ["CA-2014-112158", 41977, "Consumer", "East", "Furniture", "Bookcases", 43.92, "New York"],
        ["CA-2014-166191", 41982, "Corporate", "Central", "Technology", "Accessories", 408.744, "Illinois"],
        ["US-2014-150574", 41998, "Consumer", "South", "Office Supplies", "Binders", 4.812, "Florida"],
        ["CA-2014-106803", 42006, "Consumer", "Central", "Office Supplies", "Storage", 24.56, "Minnesota"],
        ["CA-2015-146262", 42013, "Corporate", "East", "Office Supplies", "Labels", 23.68, "Ohio"],
        ["CA-2015-110457", 42069, "Consumer", "West", "Furniture", "Tables", 787.53, "Washington"],
        ["CA-2015-145352", 42085, "Consumer", "South", "Office Supplies", "Binders", 4.95, "Georgia"],
        ["US-2015-153500", 42190, "Corporate", "East", "Office Supplies", "Paper", 26.72, "Pennsylvania"],
        ["CA-2015-155334", 42216, "Consumer", "West", "Furniture", "Furnishings", 35.28, "California"],
        ["CA-2015-144267", 42239, "Home Office", "West", "Furniture", "Chairs", 544.008, "California"],
        ["CA-2015-137946", 42251, "Consumer", "West", "Office Supplies", "Binders", 4.752, "California"],
        ["US-2015-138303", 42254, "Consumer", "East", "Office Supplies", "Storage", 36.336, "Pennsylvania"],
        ["CA-2015-130883", 42279, "Consumer", "West", "Technology", "Accessories", 239.8, "Oregon"],
        ["CA-2015-129098", 42290, "Consumer", "South", "Office Supplies", "Storage", 30.84, "Virginia"],
        ["CA-2015-102281", 42291, "Home Office", "East", "Office Supplies", "Art", 19.9, "New York"],
        ["US-2015-156867", 42325, "Consumer", "West", "Technology", "Accessories", 238.896, "Colorado"],
        ["CA-2015-135545", 42338, "Consumer", "West", "Office Supplies", "Binders", 15.824, "California"],
        ["CA-2015-160059", 42339, "Home Office", "South", "Office Supplies", "Binders", 6.24, "Arkansas"],
        ["CA-2015-101910", 42341, "Consumer", "West", "Furniture", "Chairs", 283.92, "California"],
        ["CA-2015-135272", 42350, "Consumer", "West", "Furniture", "Furnishings", 79.92, "California"],
        ["CA-2015-143490", 42351, "Consumer", "West", "Technology", "Phones", 219.184, "California"],
        ["US-2015-126214", 42362, "Consumer", "West", "Furniture", "Tables", 1618.37, "Washington"],
        ["CA-2015-117415", 42369, "Home Office", "Central", "Office Supplies", "Envelopes", 113.328, "Texas"],
        ["CA-2015-158792", 42371, "Consumer", "East", "Office Supplies", "Fasteners", 22.2, "Massachusetts"],
        ["CA-2016-126529", 42382, "Home Office", "East", "Office Supplies", "Paper", 33.312, "Ohio"],
        ["CA-2016-169103", 42442, "Consumer", "South", "Furniture", "Furnishings", 102.36, "Florida"],
        ["CA-2016-162138", 42487, "Corporate", "West", "Office Supplies", "Binders", 3.52, "California"],
        ["CA-2016-109869", 42489, "Home Office", "West", "Office Supplies", "Appliances", 78.272, "Arizona"],
        ["CA-2016-152814", 42492, "Home Office", "West", "Office Supplies", "Paper", 29.472, "Colorado"],
        ["US-2016-139486", 42513, "Consumer", "West", "Technology", "Accessories", 66.26, "California"],
        ["CA-2016-103730", 42536, "Consumer", "East", "Technology", "Phones", 68.04, "Delaware"],
        ["CA-2016-138688", 42537, "Corporate", "West", "Office Supplies", "Labels", 14.62, "California"],
        ["CA-2016-107216", 42538, "Home Office", "West", "Technology", "Accessories", 29.29, "California"],
        ["CA-2016-140081", 42545, "Home Office", "East", "Office Supplies", "Paper", 15.552, "Pennsylvania"],
        ["CA-2016-120180", 42567, "Consumer", "East", "Office Supplies", "Supplies", 11.632, "Pennsylvania"],
        ["CA-2016-110366", 42620, "Corporate", "East", "Furniture", "Furnishings", 82.8, "Pennsylvania"],
        ["CA-2016-121223", 42626, "Corporate", "East", "Office Supplies", "Paper", 8.448, "Pennsylvania"],
        ["CA-2016-161781", 42643, "Home Office", "Central", "Office Supplies", "Art", 40.88, "Indiana"],
        ["CA-2016-144939", 42651, "Consumer", "East", "Furniture", "Chairs", 599.292, "New York"],
        ["CA-2016-155516", 42664, "Corporate", "East", "Office Supplies", "Binders", 23.2, "Connecticut"],
        ["CA-2016-142545", 42677, "Corporate", "East", "Furniture", "Furnishings", 77.6, "New Jersey"],
        ["CA-2016-142545", 42677, "Corporate", "East", "Office Supplies", "Binders", 14.28, "New Jersey"],
        ["CA-2016-161669", 42683, "Corporate", "West", "Office Supplies", "Supplies", 21.36, "California"],
        ["CA-2016-128867", 42684, "Consumer", "Central", "Office Supplies", "Binders", 27.24, "Iowa"],
        ["CA-2016-105284", 42705, "Home Office", "East", "Office Supplies", "Fasteners", 4.416, "Pennsylvania"],
        ["US-2016-150861", 42710, "Consumer", "East", "Office Supplies", "Labels", 6.3, "New York"],
        ["CA-2016-161389", 42714, "Consumer", "West", "Office Supplies", "Binders", 7.976, "Washington"],
        ["CA-2017-131954", 42760, "Home Office", "West", "Office Supplies", "Binders", 27.936, "Washington"],
        ["CA-2017-127432", 42762, "Home Office", "West", "Office Supplies", "Storage", 51.45, "Montana"],
        ["CA-2017-104220", 42771, "Corporate", "Central", "Office Supplies", "Binders", 18.28, "Iowa"],
        ["CA-2017-104220", 42771, "Corporate", "Central", "Technology", "Phones", 207, "Iowa"],
        ["CA-2017-110478", 42803, "Corporate", "West", "Office Supplies", "Envelopes", 15.25, "California"],
        ["CA-2017-129567", 42815, "Consumer", "West", "Office Supplies", "Binders", 17.456, "California"],
        ["CA-2017-144932", 42842, "Consumer", "East", "Office Supplies", "Art", 89.856, "Ohio"],
        ["CA-2017-140963", 42899, "Home Office", "West", "Technology", "Phones", 279.96, "California"],
        ["CA-2017-132934", 42912, "Consumer", "East", "Office Supplies", "Binders", 5.312, "New York"],
        ["CA-2017-122105", 42914, "Consumer", "West", "Office Supplies", "Art", 95.92, "California"],
        ["CA-2017-102946", 42921, "Home Office", "West", "Office Supplies", "Binders", 6.792, "Nevada"],
        ["CA-2017-117947", 42970, "Corporate", "East", "Technology", "Phones", 37.91, "New York"],
        ["CA-2017-137596", 42985, "Home Office", "Central", "Technology", "Phones", 1199.8, "Michigan"],
        ["CA-2017-130043", 42997, "Corporate", "Central", "Office Supplies", "Paper", 31.872, "Texas"],
        ["CA-2017-133333", 43000, "Corporate", "Central", "Office Supplies", "Paper", 22.72, "Wisconsin"],
        ["CA-2017-126074", 43014, "Consumer", "Central", "Office Supplies", "Binders", 8.05, "Michigan"],
        ["CA-2017-132976", 43025, "Corporate", "East", "Office Supplies", "Paper", 11.648, "Pennsylvania"],
        ["CA-2017-132976", 43025, "Corporate", "East", "Office Supplies", "Labels", 24.84, "Pennsylvania"],
        ["CA-2017-150707", 43027, "Home Office", "East", "Office Supplies", "Binders", 37.66, "Maryland"],
        ["CA-2017-107727", 43031, "Home Office", "Central", "Office Supplies", "Paper", 29.472, "Texas"],
        ["CA-2017-125388", 43031, "Corporate", "East", "Furniture", "Furnishings", 56.56, "Massachusetts"],
        ["CA-2017-155558", 43041, "Consumer", "Central", "Office Supplies", "Labels", 6.16, "Minnesota"],
        ["CA-2017-153339", 43044, "Corporate", "South", "Furniture", "Furnishings", 15.992, "Tennessee"],
        ["US-2017-107272", 43051, "Consumer", "West", "Office Supplies", "Binders", 2.388, "Arizona"],
        ["CA-2017-150959", 43052, "Consumer", "Central", "Office Supplies", "Labels", 10.44, "Texas"],
        ["US-2017-110576", 43071, "Home Office", "East", "Furniture", "Furnishings", 36.488, "Pennsylvania"],
        ["CA-2017-145233", 43074, "Consumer", "West", "Technology", "Phones", 470.376, "Colorado"],
        ["CA-2017-117457", 43081, "Consumer", "West", "Technology", "Accessories", 179.95, "California"],
        ["CA-2017-117457", 43081, "Consumer", "West", "Office Supplies", "Paper", 27.15, "California"],
        ["US-2017-145366", 43082, "Corporate", "East", "Office Supplies", "Storage", 37.208, "Ohio"],
        ["CA-2017-155376", 43096, "Consumer", "Central", "Office Supplies", "Appliances", 839.43, "Missouri"],
      ]);

      // Prepare table data for RegionMap chart.
      let regionMapRange_State = sheet.getRange("A151:A248");
      let regionMapRange_Sales = sheet.getRange("B151:B248");
      let data1 = [];
      for (let i = 1; i < 99; i++) {
        let item = [];
        item.push("=H" + i.toString());
        data1.push(item);
      }
      regionMapRange_State.formulas = data1;
      regionMapRange_State.format.autofitColumns();
      let data2 = [];
      for (let i = 1; i < 99; i++) {
        let item = [];
        item.push("=G" + i.toString());
        data2.push(item);
      }
      regionMapRange_Sales.formulas = data2;
      regionMapRange_Sales.format.autofitColumns();

      let segment = sheet.getRange("E151:E151");
      segment.formulas = [["=SUMPRODUCT(1/COUNTIF(C2:C98,C2:C98))"]];
      let region = sheet.getRange("E152:E152");
      region.formulas = [["=SUMPRODUCT(1/COUNTIF(D2:D98,D2:D98))"]];

      let category = sheet.getRange("E153:E153");
      category.formulas = [["=SUMPRODUCT(1/COUNTIF(E2:E98,E2:E98))"]];
      let subCategory = sheet.getRange("E154:E154");
      subCategory.formulas = [["=SUMPRODUCT(1/COUNTIF(F2:F98,F2:F98))"]];
      let maxSale = sheet.getRange("E155:E155");
      maxSale.formulas = [["=MAX(G2:G98)"]];
      let sumSale = sheet.getRange("E156:E156");
      sumSale.formulas = [["=SUM(G2:G98)"]];
      let state = sheet.getRange("E157:E157");
      state.formulas = [["=SUMPRODUCT(1/COUNTIF(H2:H98,H2:H98))"]];

      await context.sync();
      showStatus('Add sample data - success!', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

function showStatus(message, isError) {
  let status = document.getElementById("status");
  // Clear previous content.
  status.innerHTML = "";

  // Create the container div.
  let statusCard = document.createElement("div");
  statusCard.className = `status-card ms-depth-4 ${isError ? "error-msg" : "success-msg"}`;

  // Create and append the first paragraph.
  let p1 = document.createElement("p");
  p1.className = "ms-fontSize-24 ms-fontWeight-bold";
  p1.textContent = isError ? "An error occurred" : "";
  statusCard.appendChild(p1);

  // Create and append the second paragraph.
  let p2 = document.createElement("p");
  p2.className = "ms-fontSize-16 ms-fontWeight-regular";
  p2.textContent = message;
  statusCard.appendChild(p2);

  // Append the status card to the status element.
  status.appendChild(statusCard);
}
