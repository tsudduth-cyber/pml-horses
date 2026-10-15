/**
 * The shape of the "PML Horses" SharePoint list.
 *
 * /api/setup creates the list from this. Nothing else reads it at runtime —
 * the app looks columns up by name — so this is the one place the schema is
 * written down, and the README table should match it.
 */

const text = () => ({ text: { allowMultipleLines: false, maxLength: 255 } });
const memo = () => ({
  text: { allowMultipleLines: true, textType: "plain", linesForEditing: 6 },
});
const dateOnly = () => ({
  dateTime: { format: "dateOnly", displayAs: "default" },
});
const money = () => ({ currency: { locale: "en-us" } });
const choice = (choices: string[]) => ({
  choice: { allowTextEntry: false, choices, displayAs: "dropDownMenu" },
});

export const LIST_COLUMNS = [
  { name: "BarnName", ...text() },
  { name: "NamePending", boolean: {} },
  { name: "IdLine", ...text() },
  { name: "AQHA", ...text() },
  {
    name: "Sex",
    ...choice(["Stallion", "Mare", "Gelding", "Filly", "Colt"]),
  },
  { name: "Color", ...text() },
  { name: "Foaled", ...dateOnly() },
  { name: "Bought", ...dateOnly() },
  { name: "BoughtNote", ...text() },
  { name: "Price", ...money() },
  { name: "SalePrice", ...money() },
  { name: "SoldOn", ...dateOnly() },
  {
    name: "Status",
    ...choice([
      "Work in progress",
      "Registered to PML",
      "Recip mare",
      "Embryo",
      "Not registering",
      "Sold",
    ]),
  },
  { name: "StatusNote", ...text() },
  { name: "Pedigree", ...text() },
  { name: "BreedingTo", ...text() },
  { name: "Notes", ...memo() },
  { name: "Missing", ...memo() },
  { name: "Farrier", ...dateOnly() },
  { name: "Deworm", ...dateOnly() },
  { name: "Vax", ...dateOnly() },
  { name: "Coggins", ...dateOnly() },
  { name: "Dental", ...dateOnly() },
  { name: "FolderName", ...text() },
  { name: "Sort", number: { decimalPlaces: "none" } },
] as const;

/** Title comes free with a generic list and holds the registered name. */
export const LIST_TEMPLATE = {
  displayName: process.env.SP_LIST_NAME ?? "PML Horses",
  list: { template: "genericList" },
  columns: LIST_COLUMNS,
};
