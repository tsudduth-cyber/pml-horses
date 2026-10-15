export type Status =
  | "wip"
  | "done"
  | "recip"
  | "exp"
  | "noreg"
  | "off";

export type CareKey = "farrier" | "worm" | "vax" | "coggins" | "dental";

export interface Doc {
  name: string;
  url: string;
  folder: boolean;
  modified: string;
}

export interface Horse {
  /** SharePoint list item id — the key for writes. */
  id: string;
  reg: string | null;
  barn: string | null;
  /** True when the horse has no name yet; the UI shows a blank, not a name. */
  pending: boolean;
  /** How to identify a pending horse — pedigree, or who she's out of. */
  idline: string | null;
  aqha: string | null;
  sex: string | null;
  color: string | null;
  dob: string | null;
  bought: string | null;
  boughtNote: string | null;
  price: number | null;
  salePrice: number | null;
  soldOn: string | null;
  status: Status;
  label: string | null;
  pedigree: string | null;
  breeding: string | null;
  notes: string[];
  missing: string[];
  care: Partial<Record<CareKey, string>>;
  folderName: string | null;
  folderUrl: string | null;
  docs: Doc[];
  photos: Doc[];
  sort: number;
}

export const STATUS_LABEL: Record<Status, string> = {
  wip: "Work in progress",
  done: "Registered to PML",
  recip: "Recip mare",
  exp: "Embryo",
  noreg: "Not registering",
  off: "Sold",
};

/**
 * Routine care intervals, in days. These are standard defaults — confirm
 * them with Philip's vet and farrier before anyone trusts a due date.
 */
export const CARE_TASKS: {
  key: CareKey;
  label: string;
  every: string;
  days: number;
}[] = [
  { key: "farrier", label: "Farrier", every: "every 6–8 weeks", days: 49 },
  { key: "worm", label: "Deworming", every: "quarterly", days: 91 },
  { key: "vax", label: "Core vaccinations", every: "yearly", days: 365 },
  { key: "coggins", label: "Coggins", every: "yearly", days: 365 },
  { key: "dental", label: "Dental", every: "yearly", days: 365 },
];

/** A live animal that could need routine care. */
export function isAnimal(h: Horse) {
  return h.status !== "off" && h.status !== "exp";
}
