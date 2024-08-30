import { useState, useEffect, useRef } from "react";
import { SetError } from "../../hooks/useErrorSnackbar";
import Storage from "../../../storage/HighlightRowStorage";
import Highlighter, { HighlighterOption } from "./Highlighter";
import Cleaner from "./HighlightCleaner";

type EventHandlerResult = OfficeExtension.EventHandlerResult<any>;

export default function useHighlightRow(setError: SetError) {
  const [isEnabled, setIsEnabled] = useState<boolean>(false);
  const [xColor, setXColor] = useState(Storage.xColor() ?? "#9CC8F5");

  const option = useRef<HighlighterOption>({ xColor });
  const eventResults = useRef<{ id: number; results: EventHandlerResult[] }>({ id: 0, results: [] });

  option.current = { xColor };

  // Cleanup function for removing event handlers
  const cleanupEventHandlers = async (results: EventHandlerResult[]) => {
    for (const event of results) {
      await Excel.run(event.context, async (context) => {
        // Remove the event handler
        event.remove();

        await context.sync();
      }).catch((error) => {
        setError([error.message, error.code, "cleanupEventHandlers"]);
      });
    }
  };

  // Initialize the class and set up event handlers
  const setEventHadlers = async (id: number) => {
    await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;

      // Register event handler for selection change
      const selection = worksheets.onSelectionChanged.add(async () => {
        await Highlighter.create(setError, option.current).updateRow();
      });

      // Register event handler for worksheet name change
      const name = worksheets.onNameChanged.add(async (event) => {
        await Storage.renameWorksheet(event.nameAfter, event.nameBefore);
      });


      await context.sync();
      const results: EventHandlerResult[] = [selection, name];

      if (eventResults.current.id === id) {
        eventResults.current.results = results;
      } else {
        cleanupEventHandlers(results);
      }
    }).catch((error) => {
      setError([error.message, error.code, "setEventHadlers"]);
    });
  }

  useEffect(() => {
    if (!isEnabled) return undefined;

    const id = performance.now();
    eventResults.current.id = id;
    setEventHadlers(id);

    return () => {
      eventResults.current.id = 0;
      cleanupEventHandlers(eventResults.current.results);
    };
  }, [isEnabled]);

  // Set the highlight color
  useEffect(() => {
    isEnabled && Highlighter.create(setError, option.current).changeColor();
    isEnabled && Storage.xColor(xColor);

    // Clear the previous highlight
    !isEnabled && Cleaner.create(setError).clearPrevious();
  }, [xColor, isEnabled]);

  return {
    isEnabled,
    setIsEnabled,
    xColor,
    setXColor,
  };
}
