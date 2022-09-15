/** Copyright (c) Microsoft Corporation. Licensed under the MIT License. */

// Set up the task pane buttons and select list.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign the HTML buttons to the relevant method.
    document.getElementById("setData").onclick = setData;
    document.getElementById("getData").onclick = getData;
    document.getElementById("clearForm").onclick = clearForm;

    // Update the displayed input fields based on the select list. 
    $("#dataTypeSelect").on("change", function() {
      let value = $("#dataTypeSelect option:selected");
      setSelectedType(getTypeContent(value.val().toString()));
    });

    setSelectedType(getTypeContent(defaultType));
  }
});

const defaultType: string = "FormattedNumber";

function textInputWithLabel(inputID: string, labelText: string, altText: string): JQuery<HTMLElement>[] {
  return [
    $("<td>").append($("<label/>", { for: inputID, class: "labels" }).text(labelText)),
    $("<td>").append($("<input/>", { class: "inputBox " + inputID, type: "text", id: inputID, alt: altText }))
  ];
}

function checkboxWithLabel(inputID: string, labelText: string, altText: string): JQuery<HTMLElement> {
  return $("<div/>", { class: "checkboxes" })
    .append($("<input/>", { type: "checkbox", class: inputID, checked: true, alt: altText }))
    .append($("<label/>").text(labelText));
}

function stringContent(): JQuery<HTMLElement> {
  return basicContent("String");
}

function booleanContent(): JQuery<HTMLElement> {
  return basicContent("Boolean");
}

function doubleContent(): JQuery<HTMLElement> {
  return basicContent("Double");
}

function basicContent(contentType: string): JQuery<HTMLElement> {
  return $("<tr/>", { id: contentType, class: "fieldValueContents" }).append(
    textInputWithLabel("basicValue", "Value: ", "value input box")
  );
}

function webImageContent(): JQuery<HTMLElement>[] {
  return [
    $("<tr/>", { id: "WebImage", class: "fieldValueContents" }).append(
      textInputWithLabel("url", "Image url: ", "Image url input box")
    ),
    $("<tr/>").append(textInputWithLabel("altText", "[Recommended] Alt-text: ", "Alt-text input box"))
  ];
}

function formattedNumberContent(): JQuery<HTMLElement>[] {
  return [
    $("<tr/>", { id: "FormattedNumber", class: "fieldValueContents" }).append(
      textInputWithLabel("number", "Number: ", "Number input box")
    ),
    $("<tr/>").append(textInputWithLabel("format", "Format: ", "Format input box"))
  ];
}

function unsupportedContent(): JQuery<HTMLElement> {
  return $("<tr/>", { id: "Unsupported", class: "fieldValueContents" })
    .append($("<td/>").append($("<label/>", { for: "unsupportedValue", class: "labels" }).text("unsupportedValue: ")))
    .append(
      $("<td/>").append(
        $("<input>", {
          class: "inputBox unsupportedValue",
          disabled: true,
          type: "text",
          name: "unsupportedValue",
          id: "unsupportedValue",
          alt: "unsupportedValue input box"
        })
      )
    );
}

function createOptionsFromList(list: string[]): JQuery<HTMLElement>[] {
  let options = [];
  list.forEach(function(val) {
    const noSpaceVal = val.replace(" ", "");
    options.push($("<option/>", { class: "type" + noSpaceVal, value: noSpaceVal }).text(val));
  });

  return options;
}

function specificFieldContent(): JQuery<HTMLElement> {
  let content = entityContents("String"); // Set the default data type inside an entity to "String".
  let label = $("<label><strong>Data type: </strong></label>");
  let select = $("<select/>", {
    id: "dataTypeSelectEntity",
    name: "dataType",
    class: "dataTypeSelectEntity ms-Button ms-Button-label buttons",
  });

  select[0].onchange = function() {
    let options = (select[0] as HTMLSelectElement).options;
    let valueType = options.item(options.selectedIndex).value;
    $(content)
      .children()
      .replaceWith(entityContents(valueType).children());
  };

  let options = createOptionsFromList(["String", "Double", "Boolean", "Web Image", "Formatted Number"]);
  for (let i = 0; i < options.length; ++i) {
    select.append(options[i]);
  }

  select.append($("<option/>", { class: "typeUnsupported", value: "Unsupported", disabled: true }).text("Unsupported"));

  let table = $("<table/>", { id: "fieldTable" }).append(
    $("<tbody/>").append($("<tr/>").append($("<td/>").append(label).append(select)))
  );

  return $("<div/>", { id: "specificField", class: "specificFieldContents formContents solidBorder" })
    .append(table)
    .append(content);
}

function entitySectionContent(): JQuery<HTMLElement> {
  let fields = $(`<div class="fields"/>`).append(specificFieldContent());

  let section = $(`<div class="collapsibleSection" aria-expanded="true"/>`).append(fields);

  let newFieldButton = $("<button/>", {
    id: "addField",
    class: "ms-Button ms-Button-label buttons",
    alt: "add another field to current section",
    style: "margin-left:20px;"
  }).text("Add another field");

  /** Add a new field to selected section in the entity contents */
  newFieldButton[0].onclick = function() {
    $(fields).append(specificFieldContent());
  };

  section.append(newFieldButton);

  return section;
}

function entityDefaultSection(): JQuery<HTMLElement> {
  let defaultSection = $(`<div id="defaultSection" class="sectionContents formContents solidBorder"/>`);
  let label = $("<label/>", { class: "sectionHeader" }).text("Default section: ");

  defaultSection.append([label, entitySectionContent()]);

  return defaultSection;
}

function entitySection(): JQuery<HTMLElement> {
  let element = $("<div/>", { class: "sectionContents formContents solidBorder" });
  let table1 = $(`
              <table id="sectionTable">
                <tbody>
                  <tr class="columnTitle">
                    <td><label class="sectionHeader">Section Title:</label></td>
                    <td><input class="inputBox sectionTitle" alt="section title input box"/></td>
                    <td>
                      <button class="ms-Button ms-Button-label arrows sectionToggle" onclick="collapseSection(this)" alt="collapse section" title="collapse section" style="visibility:visible">&#x2228</button>
                    </td>
                  </tr>
                </tbody>
              </table>`);

  let removeSectionButton = $("<button/>", {
    class: "ms-Button ms-Button-label buttons",
    alt: "delete current section and its contents"
  }).text("Delete Section and its Contents");

  /** Remove the selected section in the entity contents. */
  removeSectionButton[0].onclick = function() {
    $(element).remove();
  };

  element.append(table1);
  element.append(entitySectionContent().append(removeSectionButton));

  return element;
}

function getTypeContent(valueType: string): JQuery<HTMLElement> | JQuery<HTMLElement>[] {
  switch (valueType) {
    case "String":
      return stringContent();
    case "Double":
      return doubleContent();
    case "Boolean":
      return booleanContent();
    case "FormattedNumber":
      return formattedNumberContent();
    case "WebImage":
      return webImageContent();
    case "Entity":
      return entityContent();
    case "Unsupported":
    default:
      return unsupportedContent();
  }
}

/** Create the HTML for entity contents section. */
function entityContents(valueType: string): JQuery<HTMLElement> {
  let trKey = $(`<tr>
          <td><label class="labels">Key:</label></td>
          <td><input class="inputBox fieldName" alt="Key input box"/></td>
        </tr>`);

  let tdMetadata = $(`<td colspan="4" class="center settings"/>`);

  if (valueType == "WebImage") {
    tdMetadata.append(checkboxWithLabel("mainImage", "Make main image", "main image checkbox"));
    tdMetadata.append($("<br/>"));
  }

  tdMetadata.append([
    checkboxWithLabel("cardView", "Card View", "cardview checkbox"),
    checkboxWithLabel("autoComplete", "Autocomplete", "autoComplete checkbox"),
    $("<br/>"),
    checkboxWithLabel("calcCompare", "Calc compare", "calcCompare checkbox"),
    checkboxWithLabel("dotNotation", "Dot notation", "dotNotation checkbox"),
    $("<br/>"),
    $("<label/>").text("Sublabel: "),
    $("<input/>", { class: "sublabel", alt: "sublabel input box" })
  ]);

  let trMetadata = $(`<tr class="metadata" style="visibility:collapse"/>`).append(tdMetadata);

  let trButtons = $(`<tr/>`);
  let tdButtons = $(`<td colspan="4"/>`);

  let buttonToggleMetadata = $(
    `<button class="ms-Button ms-Button-label buttons" alt="Toggle to expand or collapse metadata properties of field">More settings</button>`
  );

  /** Expand or collapse the additional metadata contents of a
  particular input field, within the entity contents. */
  buttonToggleMetadata[0].onclick = function() {
    let visibility = trMetadata[0].style.visibility;
    if (visibility != "collapse") {
      trMetadata[0].style.visibility = "collapse";
    } else {
      trMetadata[0].style.visibility = "visible";
    }
  };

  tdButtons.append([buttonToggleMetadata]);
  trButtons.append(tdButtons);

  return $(`<div id="entityContents"/>`).append(
    $(`<table id="fieldTable"/>`).append(
      $(`<tbody/>`)
        .append(trKey)
        .append(getTypeContent(valueType))
        .append([trMetadata, trButtons])
    )
  );
}

/** Create the HTML for the entity data type. */
function entityContent(): JQuery<HTMLElement> {
  let div = $(`<div class="contentPadding"/>`);
  let iconlabel = $(`<label class="labels"> Entity icon: </label>`);
  let contentLabel = $(`<label class="labels contentPadding">Entity contents:</label>`);
  let select = $(`<select id="iconSelect" name="dataType" class="ms-Button ms-Button-label buttons"/>`);
  let options = createOptionsFromList([
    "Generic",
    "Airplane",
    "Animal",
    "Apple",
    "Art",
    "Atom",
    "Bank",
    "Basketball",
    "Beaker",
    "Bird",
    "Book",
    "Bridge",
    "Briefcase",
    "Car",
    "Cat",
    "City",
    "Clouds",
    "Constellation",
    "Dinosaur",
    "Disaster",
    "DNA",
    "Dog",
    "Drama",
    "Galaxy",
    "HatGraduation",
    "Heart",
    "Languages",
    "Leaf",
    "Location",
    "Map",
    "Microscope",
    "Money",
    "Mountain",
    "MovieCamera",
    "MusicNote",
    "Notebook",
    "PartlySunnyWeather",
    "Person",
    "Planet",
    "PointScan",
    "Running",
    "Satellite",
    "Syringe",
    "Violin",
    "Wand"
  ]);

  for (let i = 0; i < options.length; ++i) {
    select.append(options[i]);
  }

  div.append(iconlabel);
  div.append(select);

  let sections = $(`<div class="sections"/>`).append(entityDefaultSection());

  let label = $(`<label for="displayString" class="labels">Entity display text: </label>`);
  let input = $(`<input class="inputBox displayString" type="text" id="displayString" alt="display text input box"/>`);

  let providerTable = $(`<div id="entityContents" class="solidBorder formContents"/>`).append(
    $(`<table id="fieldTable"/>`).append(
      $(`<tbody/>`)
        .append(
          $("<tr/>").append(
            textInputWithLabel("providerDescription", "Description: ", "Provider description input box")
          )
        )
        .append(
          $("<tr/>").append(textInputWithLabel("providerTarget", "Provider address: ", "Provider source input box"))
        )
        .append($("<tr/>").append(textInputWithLabel("providerLogo", "Logo address: ", "Provider logo input box")))
    )
  );

  let referencedValuesLabel = $(`<label for="referencedValues" class="labels">Referenced values: </label>`);
  let referencedValuesInput = $(
    `<input class="inputBox displayString" type="text" id="referencedValues" alt="referencedValues text display box" disabled/>`
  );
  let referencedValuesClearButton = $("<button/>", {
    id: "clearReferencedValue",
    class: "ms-Button ms-Button-label buttons",
    text: "Clear referencedValues"
  });

  let referencedValuesDiv = $("<div/>", { id: "referencedValuesDiv" })
    .append(referencedValuesLabel)
    .append(referencedValuesInput)
    .append(referencedValuesClearButton);

  referencedValuesDiv[0].style.display = "none";

  referencedValuesClearButton[0].onclick = function() {
    referencedValuesInput.val("");
    referencedValuesDiv[0].style.display = "none";
  };

  let providerInfo = $("<div/>", { class: "contentPadding" })
    .append($("<label/>", { class: "labels" }).text("Provider info: "))
    .append(providerTable);

  let element = $(`<tr id="Entity"/>`).append(
    $(`<td colspan="2"/>`)
      .append([label, input])
      .append(div)
      .append(providerInfo)
      .append(referencedValuesDiv)
      .append(contentLabel)
      .append(sections)
  );

  return element;
}

/** Create the HTML for when boolean, string, or double data types are selected */
function setSelectedType(selected: JQuery<HTMLElement>[] | JQuery<HTMLElement>) {
  let element = $(`<div class= "backgroundColorForm solidBorder"/>`);
  element.append(selected);

  $(".backgroundColorForm").replaceWith(element);
}

/** Add a new field to the entity fields. */
function addField(element: HTMLButtonElement) {
  // Append specificFieldContents to the fieldsDiv.
  const fieldsDiv = element.parentElement.previousElementSibling;
  $(fieldsDiv).append(specificFieldContent());
}

/** Expand the section fields in the entity field. */
function expandSection(element: HTMLButtonElement) {
  $(element).replaceWith(
    `<button class="ms-Button ms-Button-label arrows sectionToggle" onclick="collapseSection(this)" alt="collapse section" title="collapse section">&#x2228</button>`
  );
}

/** Collapse the section fields in the entity field. */
function collapseSection(element: HTMLButtonElement) {
  $(element).replaceWith(
    `<button class="ms-Button ms-Button-label arrows sectionToggle" onclick="expandSection(this)" alt="expand section" title="expand section" >&#x2227 </button>`
  );
}

/** Assign the inputted data to the active worksheet cell as the appropriate data type. */
function createValueAsJson(): Excel.CellValue {
  let values = $("#dataTypeSelect option:selected");
  switch (values.val()) {
    case "String":
      let stringValue = $("#basicValue")
        .val()
        .toString();
      return {
        type: Excel.CellValueType.string,
        basicValue: stringValue
      } as Excel.StringCellValue;

    case "Double":
      var doubleValue = Number($("#basicValue").val());
      if (!isNaN(doubleValue)) {
        return {
          type: Excel.CellValueType.double,
          basicValue: doubleValue
        } as Excel.DoubleCellValue;
      } else {
        alert("Type 'Double' selected but input was not a double.");
      }
      break;

    case "Boolean":
      let booleanValue = $("#basicValue")
        .val()
        .toString();
      if (booleanValue.toLowerCase() === "true") {
        return {
          type: Excel.CellValueType.boolean,
          basicValue: true
        } as Excel.BooleanCellValue;
      } else if (booleanValue.toLowerCase() === "false") {
        return {
          type: Excel.CellValueType.boolean,
          basicValue: false
        } as Excel.BooleanCellValue;
      } else {
        alert("Type 'Boolean' selected but input was not a boolean.");
      }
      break;

    case "Entity":
      return setEntity();

    case "WebImage":
      let url = $("#url")
        .val()
        .toString();
      let altText = $("#altText")
        .val()
        .toString();
      return {
        type: Excel.CellValueType.webImage,
        address: url,
        altText: altText
      };

    case "FormattedNumber":
      var doubleValue = Number($("#number").val());
      let format = $("#format").val();
      if (!isNaN(doubleValue)) {
        return {
          type: Excel.CellValueType.formattedNumber,
          basicValue: doubleValue,
          numberFormat: format
        } as Excel.FormattedNumberCellValue;
      } else {
        alert("Type 'FormattedNumber' selected but input was not a number.");
      }
      break;
  }
}

async function setData() {
  await Excel.run(async (context) => {
    const activeCell = context.workbook.getActiveCell();
    activeCell.valuesAsJson = [[createValueAsJson()]];
    await tryCatch(context.sync);
  });
}

/** Assign the inputted entity contents to an entity data type. */
function setEntity() {
  const display: string = $("#displayString")
    .val()
    .toString();
  let iconName: string = $("#iconSelect option:selected").val() as string;
  const referencedValues: string = $("#referencedValues").val() as string;
  const fields = valuesFromQuery(".fieldName");
  const values = fieldValuesContentsFromQuery();
  const cardViews = valuesFromQuery(".cardView");
  const autoCompletes = valuesFromQuery(".autoComplete");
  const calcCompares = valuesFromQuery(".calcCompare");
  const dotNotation = valuesFromQuery(".dotNotation");
  const sublabels = valuesFromQuery(".sublabel");
  let mainImage = valuesFromQuery(".mainImage");
  let providerInfo = [
    valuesFromQuery("#providerDescription"),
    valuesFromQuery("#providerTarget"),
    valuesFromQuery("#providerLogo")
  ];
  let mainImageExists = false;
  let mainImageKey;
  let sectionArray = [];
  let jqSectionContents = $(".sectionContents");

  let fDefaultSection: Boolean = true;
  while (jqSectionContents.length > 0) {
    const first = jqSectionContents.first();
    let children = first.find(".fieldName");
    let sectionTitle = first.find(".sectionTitle");
    let properties = [];
    for (let i = 0; i < children.length; ++i) {
      let val = $(children[i]).val();
      properties.push(val);
    }

    // First section is the default; not really a section.
    if (fDefaultSection) {
      fDefaultSection = false;
      jqSectionContents = jqSectionContents.slice(1);
      continue;
    }

    let sectionEntry = {
      layout: "List",
      title: sectionTitle.val(),
      properties: properties
    };
    sectionArray.push(sectionEntry);
    jqSectionContents = jqSectionContents.slice(1);
  }

  let entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: display,
    properties: {},
    layouts: {
      card: {},
      compact: {}
    }
  };

  if (providerInfo[0][0] != "" || providerInfo[1][0] != "" || providerInfo[2][0] != "") {
    entity.provider = {
      description: providerInfo[0][0],
      logoTargetAddress: providerInfo[1][0],
      logoSourceAddress: providerInfo[2][0]
    };
  }

  for (let i = 0; i < fields.length; ++i) {
    let curSectionFields = [];
    const field = fields[i];
    let value = values[i];
    if (field == "" || value == "") {
      break;
    }
    let featureIntegration = {};
    if (!cardViews[i]) {
      featureIntegration["cardView"] = true;
    }
    if (!autoCompletes[i]) {
      featureIntegration["autoComplete"] = true;
    }
    if (!calcCompares[i]) {
      featureIntegration["calcCompare"] = true;
    }
    if (!dotNotation[i]) {
      featureIntegration["dotNotation"] = true;
    }
    let propertyMetadata = {};
    if (Object.keys(featureIntegration).length > 0) {
      propertyMetadata["excludeFrom"] = featureIntegration;
    }
    if (sublabels[i] != "") {
      propertyMetadata["sublabel"] = sublabels[i];
    }
    if (value.type == "WebImage") {
      if (mainImage[0] && !mainImageExists) {
        mainImageKey = field;
        mainImageExists = true;
      }
      mainImage = mainImage.slice(1);
    }

    if (Object.keys(propertyMetadata).length > 0) {
      value["propertyMetadata"] = propertyMetadata;
    }
    entity.properties[field] = value;
  }
  if (mainImageExists) {
    entity.layouts.card = {
      mainImage: {
        property: mainImageKey
      },
      sections: sectionArray
    };
  } else {
    entity.layouts.card = {
      sections: sectionArray
    };
  }

  if (referencedValues != "") {
    entity.referencedValues = JSON.parse(referencedValues);
  }

  if (iconName != "Generic") {
    entity.layouts.compact["icon"] = iconName;
  }
  return entity;
}

/** Helper function to retrieve jQuery values for setEntity(). */
function valuesFromQuery(query: string) {
  let jq = $(query);
  let result = [];
  while (jq.length > 0) {
    const first = jq.first();
    if (first.is("input[type=checkbox]")) {
      result.push(first.prop("checked"));
    } else {
      result.push(first.val());
    }
    jq = jq.slice(1);
  }
  return result;
}
/** Helper function to retrieve fieldValue contents for setEntity(). */
function fieldValuesContentsFromQuery() {
  let jqFieldContents = $(".fieldValueContents");
  let jqBasicValue = $(".basicValue");
  let jqUrl = $(".url");
  let jqNumber = $(".number");
  let jqFormat = $(".format");
  let jqAltText = $(".altText");
  let jqUnsupportedValue = $(".unsupportedValue");
  let values = [];
  while (jqFieldContents.length > 0) {
    const first = jqFieldContents.first();
    let valueType = first.attr("id");
    let value;
    switch (valueType) {
      case "String":
        value = {
          type: Excel.CellValueType.string,
          basicValue: jqBasicValue
            .first()
            .val()
            .toString()
        } as Excel.StringCellValue;
        jqBasicValue = jqBasicValue.slice(1);
        break;
      case "FormattedNumber":
        var doubleValue = Number(jqNumber.first().val());
        if (!isNaN(doubleValue)) {
          value = {
            type: Excel.CellValueType.formattedNumber,
            basicValue: doubleValue,
            numberFormat: jqFormat
              .first()
              .val()
              .toString()
          };
        } else {
          alert("Type 'FormattedNumber' selected but input was not a number.");
        }
        jqNumber = jqNumber.slice(1);
        jqFormat = jqFormat.slice(1);
        break;
      case "Double":
        var doubleValue = Number(jqBasicValue.first().val());
        if (!isNaN(doubleValue)) {
          value = {
            type: Excel.CellValueType.double,
            basicValue: doubleValue
          } as Excel.DoubleCellValue;
        } else {
          alert("Type 'Double' selected but input was not a double.");
        }
        jqBasicValue = jqBasicValue.slice(1);
        break;
      case "Boolean":
        let booleanValue = jqBasicValue
          .first()
          .val()
          .toString();
        if (booleanValue.toLowerCase() === "true") {
          value = {
            type: Excel.CellValueType.boolean,
            basicValue: true
          } as Excel.BooleanCellValue;
        } else if (booleanValue.toLowerCase() === "false") {
          value = {
            type: Excel.CellValueType.boolean,
            basicValue: false
          } as Excel.BooleanCellValue;
        } else {
          alert("Type 'Boolean' selected but input was not a boolean.");
        }
        jqBasicValue = jqBasicValue.slice(1);
        break;
      case "WebImage":
        value = {
          type: Excel.CellValueType.webImage,
          address: jqUrl.first().val(),
          altText: jqAltText.first().val()
        } as Excel.WebImageCellValue;
        jqUrl = jqUrl.slice(1);
        jqAltText = jqAltText.slice(1);
        break;
      case "Unsupported":
        value = JSON.parse(
          jqUnsupportedValue
            .first()
            .val()
            .toString()
        );
        jqUnsupportedValue = jqUnsupportedValue.slice(1);
        break;
    }
    values.push(value);
    jqFieldContents = jqFieldContents.slice(1);
  }

  return values;
}

/** Retrieve the contents of a selected entity and put them in the form boxes. */
function getEntity(value) {
  let sections;
  if (value.layouts != undefined && value.layouts.card != undefined) {
    sections = value.layouts.card.sections;
  }

  // If there is no section, make one.
  if (sections == undefined) {
    const propertyKeys = Object.keys(value.properties);
    sections = [
      {
        layout: "List",
        properties: propertyKeys
      }
    ];
  }

  if (value.referencedValues != undefined) {
    $("#referencedValues").val(JSON.stringify(value.referencedValues));
    $("#referencedValuesDiv")[0].style.display = "block";
  } else {
    $("#clearReferencedValue")[0].onclick(null);
  }

  if (value.provider != undefined) {
    if (value.provider.description != undefined) {
      $("#providerDescription").val(value.provider.description);
    }

    if (value.provider.description != undefined) {
      $("#providerTarget").val(value.provider.logoTargetAddress);
    }

    if (value.provider.description != undefined) {
      $("#providerLogo").val(value.provider.logoSourceAddress);
    }
  }

  let propertyKeysOrdered = [];
  for (let i = 0; i < sections.length; ++i) {
    $(".sectionTitle")
      .last()
      .val(sections[i].title);
    let sectionKeys = sections[i].properties;
    $(".specificFieldContents")
      .last()
      .remove();
    for (let j = 0; j < sectionKeys.length; ++j) {
      $(".fields")
        .last()
        .append(specificFieldContent());
      const propertyName = sectionKeys[j];
      propertyKeysOrdered.push(propertyName);
      const propertyValue = value.properties[propertyName];
      const entityContentsDiv = $(".dataTypeSelectEntity")
        .last()
        .parent()
        .parent()
        .parent()
        .parent()
        .next()
        .children();
      switch (propertyValue.type) {
        case "Double":
          $(".dataTypeSelectEntity")
            .last()
            .val("Double");
          entityContentsDiv.replaceWith(entityContents(propertyValue.type).children());
          break;
        case "Boolean":
          $(".dataTypeSelectEntity")
            .last()
            .val("Boolean");
          entityContentsDiv.replaceWith(entityContents(propertyValue.type).children());
          break;
        case "FormattedNumber":
          $(".dataTypeSelectEntity")
            .last()
            .val("FormattedNumber");
          entityContentsDiv.replaceWith(entityContents(propertyValue.type).children());
          break;
        case "WebImage":
          $(".dataTypeSelectEntity")
            .last()
            .val("WebImage");
          entityContentsDiv.replaceWith(entityContents(propertyValue.type).children());
          const firstMetadataElement = $(".cardView")
            .last()
            .parent();
          firstMetadataElement.before(`<div class="checkboxes">
                  <input type="checkbox" class="mainImage" alt="main image checkbox"/>
                  <label>Make main image</label>
                </div><br>`);
          break;
        case "String":
          $(".dataTypeSelectEntity")
            .last()
            .val("String");
          entityContentsDiv.replaceWith(entityContents(propertyValue.type).children());
          break;
        default:
          /*unsupported*/
          $(".dataTypeSelectEntity")
            .last()
            .val("Unsupported");
          entityContentsDiv.replaceWith(entityContents("Unsupported").children());
          break;
      }
    }
  }

  // Write the entity's data into the table.
  $("#displayString").val(value.text);
  if (value.layouts == undefined || value.layouts.compact == undefined) $("#iconSelect").val("Generic");
  else if (value.layouts.compact["icon"] != undefined) {
    $("#iconSelect").val(value.layouts.compact["icon"]);
  }
  let jqFields = $(".fieldName");
  let jqCardView = $(".cardView");
  let jqAutoComplete = $(".autoComplete");
  let jqCalcCompare = $(".calcCompare");
  let jqDotNotation = $(".dotNotation");
  let jqSublabel = $(".sublabel");
  let jqBasicValue = $(".basicValue");
  let jqUrl = $(".url");
  let jqNumber = $(".number");
  let jqFormat = $(".format");
  let jqAltText = $(".altText");
  let jqMainImage = $(".mainImage");
  let jqUnsupportedValue = $(".unsupportedValue");

  for (let i = 0; i < propertyKeysOrdered.length; ++i) {
    const propertyName = propertyKeysOrdered[i];
    const propertyValue = value.properties[propertyName];
    jqFields.first().val(propertyName);
    switch (propertyValue.type) {
      case "Double":
      case "Boolean":
      case "String":
        jqBasicValue.first().val(propertyValue.basicValue);
        jqBasicValue = jqBasicValue.slice(1);
        break;
      case "WebImage":
        jqUrl.first().val(propertyValue.address);
        jqAltText.first().val(propertyValue.altText);
        jqUrl = jqUrl.slice(1);
        jqAltText = jqAltText.slice(1);
        if (value.layouts != undefined)
          if (value.layouts.card["mainImage"] != undefined) {
            if (value.layouts.card.mainImage.property == propertyName) {
              jqMainImage.first().prop("checked", true);
            }
          }
        break;
      case "FormattedNumber":
        jqNumber.first().val(propertyValue.basicValue);
        jqFormat.first().val(propertyValue.numberFormat);
        jqNumber = jqNumber.slice(1);
        jqFormat = jqFormat.slice(1);
        break;
      default:
        jqUnsupportedValue.first().val(JSON.stringify(propertyValue));
        jqUnsupportedValue = jqUnsupportedValue.slice(1);
        break;
    }
    let featureIntegration: Excel.CellValuePropertyMetadataExclusions = {
      cardView: false,
      autoComplete: false,
      calcCompare: false,
      dotNotation: false
    };
    let sublabel = "";
    if (typeof propertyValue.propertyMetadata == "object") {
      if (typeof propertyValue.propertyMetadata.excludeFrom == "object") {
        featureIntegration = Object.assign(featureIntegration, propertyValue.propertyMetadata.excludeFrom);
      }
      if (typeof propertyValue.propertyMetadata.sublabel == "string") {
        sublabel = propertyValue.propertyMetadata.sublabel;
      }
    }
    jqCardView.first().prop("checked", !featureIntegration.cardView);
    jqAutoComplete.first().prop("checked", !featureIntegration.autoComplete);
    jqCalcCompare.first().prop("checked", !featureIntegration.calcCompare);
    jqDotNotation.first().prop("checked", !featureIntegration.dotNotation);
    jqSublabel.first().val(sublabel);
    jqFields = jqFields.slice(1);
    jqCardView = jqCardView.slice(1);
    jqAutoComplete = jqAutoComplete.slice(1);
    jqCalcCompare = jqCalcCompare.slice(1);
    jqDotNotation = jqDotNotation.slice(1);
    jqSublabel = jqSublabel.slice(1);
  }
}

/** Retrieve the contents of a selected cell and put them in the form boxes. */
async function getData() {
  await Excel.run(async (context) => {
    const activeCell = context.workbook.getActiveCell();
    activeCell.load("valuesAsJson");
    await context.sync();
    const value = activeCell.valuesAsJson[0][0];
    clearForm();

    let valueType = value.type == "LinkedEntity" ? "Entity" : value.type;
    $("#dataTypeSelect").val(valueType);
    setSelectedType(getTypeContent(valueType));

    switch (value.type) {
      case "String":
        $("#basicValue").val(value.basicValue);
        break;
      case "Double":
        $("#basicValue").val(value.basicValue);
        break;
      case "Boolean":
        const basicValue = value.basicValue;
        if (basicValue) {
          $("#basicValue").val("true");
        } else {
          $("#basicValue").val("false");
        }
        break;
      case "WebImage":
        $("#url").val(value.address);
        $("#altText").val(value.altText);
        break;
      case "FormattedNumber":
        $("#format").val(value.numberFormat);
        $("#number").val(value.basicValue);
        break;
      case "Entity":
      case "LinkedEntity":
        getEntity(value);
        break;
    }
  });
}

/** Clear the input boxes. */
async function clearForm() {
  $(".inputBox").val("");
  $(".cardView").prop("checked", true);
  $(".autoComplete").prop("checked", true);
  $(".calcCompare").prop("checked", true);
  $(".dotNotation").prop("checked", true);
  $(".mainImage").prop("checked", false);
  $(".sublabel").val("");
  $(".specificFieldContents").remove();
  $(".sectionContents").remove();
  $("#iconSelect").val("Generic");

  setSelectedType(
    getTypeContent(
      $("#dataTypeSelect")
        .val()
        .toString()
    )
  );
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    alert("Error in running script:\n\n" + error + ".");
    console.error(error);
  }
}
