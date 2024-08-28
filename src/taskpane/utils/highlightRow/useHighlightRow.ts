import { useState, useEffect, useRef } from "react";
import { SetError } from "../../hooks/useErrorSnackbar";
import Storage from "../../../storage/HighlightRowStorage";
import Highlighter, { HighlighterOption } from "./Highlighter";
import Cleaner from "./HighlightCleaner";

type EventHandlerResult = OfficeExtension.EventHandlerResult<any>;

export default function useHighlightRow(setError: SetError) {
  const [isEnabled, setIsEnabled] = useState<boolean>(false);
  const [keepSelection, setKeepSelection] = useState<boolean>(false);
  const [xColor, setXColor] = useState(Storage.xColor() ?? "#9CC8F5");

  const option = useRef<HighlighterOption>({ keepSelection, xColor });
  const eventResults = useRef<EventHandlerResult[]>([]);

  option.current = { keepSelection, xColor };

  // Initialize the class and set up event handlers
  const setEventHadlers = async () => {
    await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;

      // Register event handler for selection change
      const selection = worksheets.onSelectionChanged.add(async (event) => {
        // Check if the selection is a single cell when the keepSelection option is enabled
        if(!event.address.includes(":") && option.current.keepSelection) return;
        
        await Highlighter.create(setError, option.current).updateRow();
      });

      // Register event handler for single click
      const click = worksheets.onSingleClicked.add(async () => {
        // Check if the keepSelection option is enabled
        if (!option.current.keepSelection) return;
        await Highlighter.create(setError, option.current).updateRow();
      });

      // Register event handler for worksheet name change
      const name = worksheets.onNameChanged.add(async (event) => {
        await Storage.renameWorksheet(event.nameAfter, event.nameBefore);
      });

      await context.sync();
      eventResults.current.push(selection, name, click);
    }).catch((error) => {
      this.setError([error.message, error.code, "setEventHadlers"]);
    });
  }

  // Cleanup function for removing event handlers
  const cleanupEventHandlers = async () => {
    for (const event of eventResults.current) {
      await Excel.run(event.context, async (context) => {
        // Remove the event handler
        event.remove();

        await context.sync();
      }).catch((error) => {
        this.setError([error.message, error.code, "cleanupEventHandlers"]);
      });
    }

    eventResults.current = [];
  };

  useEffect(() => {
    isEnabled && setEventHadlers()

    return () => {
      isEnabled && cleanupEventHandlers();
    };
  }, [isEnabled]);

  // Set the highlight color
  useEffect(() => {
    isEnabled && Highlighter.create(setError, option.current).changeColor();
    isEnabled && Storage.xColor(xColor);

    if (!keepSelection && !isEnabled) {
      // Clear the previous conditional format
      Cleaner.create(setError).clearPrevious();
    }
  }, [xColor, isEnabled, keepSelection]);

  return {
    isEnabled,
    setIsEnabled,
    xColor,
    setXColor,
    keepSelection,
    setKeepSelection
  };
}
