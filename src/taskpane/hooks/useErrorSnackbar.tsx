import { useReducer } from "react";

const initialState: string[] = [];
const reducer = (state: string[], action: string[]) => {
  if (!action.length) return [];

  switch (action[1] ?? "") {
    case "InvalidOperationInCellEditMode":
      return state;
    default:
      return [...state, ...action];
  }
};

/**
 * Adds the strings in the array to the message stack. If an empty array is added, the entire state stack will be empty.
 */
export type SetError = React.Dispatch<string[]>;

/**
 *
 * @returns A tuple containing the current message stack and a function to add messages to the stack.
 */
export default function useErrorSnackbar(): [string[], SetError] {
  return useReducer(reducer, initialState);
}
