import { useState, useEffect, useCallback } from "react";
import { SetError } from "../../hooks/useErrorSnackbar";
import Profile from "./ColomnProfile";
import Storage from "../../../storage/ColumnStorage";

/**
 * Hook to manage the column width profile
 */
export default function useColumnProfile(setError: SetError, rangeAddress: string = "A1:Z1") {
  const [activeSheetName, setActiveSheetName] = useState("");
  const [profiles, setProfiles] = useState<Storage.Profiles[keyof Storage.Profiles]>([]);
  const [previous, setPrevious] = useState<{ [key: string]: number[] }>({});

  useEffect(() => {
    Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;

      // Register event handler for worksheet activation
      worksheets.onActivated.add(async () => {
        await Profile.getActiveSheetName(setError, setActiveSheetName);
      });

      // Register event handler for worksheet rename
      worksheets.onNameChanged.add(async (event) => {
        await Storage.renameWorksheet(event.nameAfter, event.nameBefore);
        await Profile.getActiveSheetName(setError, setActiveSheetName);
      });
    });

    Profile.getActiveSheetName(setError, setActiveSheetName);
  }, []);

  // Get the column profile from the storage
  useEffect(() => {
    activeSheetName && setProfiles(Storage.getProfile(activeSheetName));
  }, [activeSheetName]);

  // Add the column profile to the storage
  const add = useCallback(async () => {
    await Profile.getColumnsFromWorksheet(rangeAddress, setError, async (activeSheetName, columns) => {
      await Storage.addProfile(activeSheetName, columns);
      setProfiles(Storage.getProfile(activeSheetName));
    });
  }, [rangeAddress]);

  // Restore the column profile
  const restore = useCallback(
    async (index: number) => {
      await Profile.getColumnsFromWorksheet(rangeAddress, setError, async (activeSheetName, columns) => {
        setPrevious({ ...previous, [activeSheetName]: columns });
        await Profile.setColumnsToWorksheet(profiles[index], rangeAddress, setError);
      });
    },
    [profiles, rangeAddress, previous]
  );

  // Remove the column profile from the storage
  const remove = useCallback(
    async (index: number) => {
      await Storage.deleteProfile(activeSheetName, index);
      setProfiles(Storage.getProfile(activeSheetName));
    },
    [activeSheetName]
  );

  // Undo the column profile
  const undo = useCallback(async () => {
    if (!previous[activeSheetName]?.length) return;

    await Profile.setColumnsToWorksheet(previous[activeSheetName], rangeAddress, setError, () => {
      setPrevious({ ...previous, [activeSheetName]: [] });
    });
  }, [activeSheetName, rangeAddress, previous]);

  return {
    add,
    restore,
    remove,
    profilesCount: profiles.length,
    undo,
    hasPrevious: !!previous[activeSheetName]?.length,
  };
}
