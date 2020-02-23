/**
 * Helper functions
 * 
 * Author: Eric Dong
 * Creation Date: 2/23/20
 * Last Modfied: 2/23/20
 * 
 */

export const getValueFromColumnName = (name, row, ColumnNames) => {
  const index = ColumnNames.indexOf(name);
  if (index !== -1) {
      return row[index];
  }
  return undefined;
};