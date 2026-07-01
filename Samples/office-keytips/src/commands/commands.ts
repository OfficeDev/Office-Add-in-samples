import { insertBlueParagraphInWord } from "./word";
import { setRangeColorInExcel } from "./excel";
import { insertTextInPowerPoint } from "./powerpoint";

/* global Office console */

/**
 * Creates a named action handler that logs which ribbon control invoked it
 * and runs the host-appropriate sample action using the supplied color so
 * each control produces a distinct result.
 */
function makeAction(controlId: string, color: string) {
  return async function (event: Office.AddinCommands.Event) {
    console.log(`KeyTip control invoked: ${controlId} (color: ${color})`);
    switch (Office.context.host) {
      case Office.HostType.Excel:
        await setRangeColorInExcel(event, color);
        break;
      case Office.HostType.Word:
        await insertBlueParagraphInWord(event, color);
        break;
      case Office.HostType.PowerPoint:
        await insertTextInPowerPoint(event, color);
        break;
      default:
        event.completed();
    }
  };
}

// Register the add-in commands with the Office host application.
Office.onReady(async (info) => {
  // Original action used by the Home tab button.
  switch (info.host) {
    case Office.HostType.Word:
      Office.actions.associate("action", insertBlueParagraphInWord);
      break;
    case Office.HostType.Excel:
      Office.actions.associate("action", (event: Office.AddinCommands.Event) =>
        setRangeColorInExcel(event, "yellow")
      );
      break;
    case Office.HostType.PowerPoint:
      Office.actions.associate("action", insertTextInPowerPoint);
      break;
    default:
      throw new Error(`${info.host} not supported.`);
  }

  // Custom tab buttons. Each applies a unique color.
  Office.actions.associate("btn1Action", makeAction("Btn1", "red"));
  Office.actions.associate("btn2Action", makeAction("Btn2", "orange"));
  Office.actions.associate("btn3Action", makeAction("Btn3", "yellow"));
  Office.actions.associate("btn4Action", makeAction("Btn4", "green"));
  Office.actions.associate("btn5Action", makeAction("Btn5", "blue"));
  Office.actions.associate("btn6Action", makeAction("Btn6", "purple"));

  // Custom tab menu items. Each applies a unique color.
  Office.actions.associate("menuItem1Action", makeAction("MenuItem1", "pink"));
  Office.actions.associate("menuItem2Action", makeAction("MenuItem2", "cyan"));
  Office.actions.associate("menuItem3Action", makeAction("MenuItem3", "magenta"));
  Office.actions.associate("menuItem4Action", makeAction("MenuItem4", "lime"));
  Office.actions.associate("menuItem5Action", makeAction("MenuItem5", "brown"));
  Office.actions.associate("menuItem6Action", makeAction("MenuItem6", "gray"));
});
