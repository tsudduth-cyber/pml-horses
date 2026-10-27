export type Status =
  | "wip"
  | "done"
  | "recip"
  | "exp"
  | "noreg"
  | "off";

export type CareKey = "farrier" | "worm" | "vax" | "coggins" | "dental";

/** A care date read off a filed document rather than entered by hand. */
export interface DerivedCare {
  date: string;
  /** The file it came from, so the number is never unexplained. */
  from: string;
  /** True when it came from the barn-wide folder, not this horse's. */
  barn: boolean;
}

export interface Doc {
  /** Drive item id. Photos are fetched through /api/photo by this. */
  id: string;
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
  /** Entered by hand in the app. Always wins over a derived date. */
  care: Partial<Record<CareKey, string>>;
  /** Read off filed paperwork. Used only where `care` has nothing. */
  derived: Partial<Record<CareKey, DerivedCare>>;
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
  /** Words in a filename that mark it as paperwork for this task. */
  match: RegExp;
}[] = [
  {
    key: "farrier",
    label: "Farrier",
    every: "every 6–8 weeks",
    days: 49,
    // Half sets, full sets and pads are shoeing too — the farrier's own
    // invoice wording, which is what these have to match.
    match: /\b(farrier|shoe\w*|trim\w*|reset|half\s?set|full\s?set|pads?|clips?)\b/i,
  },
  {
    key: "worm",
    label: "Deworming",
    every: "quarterly",
    days: 91,
    match: /\b(deworm\w*|worming|wormer|worm)\b/i,
  },
  {
    key: "vax",
    label: "Core vaccinations",
    every: "yearly",
    days: 365,
    match: /\b(vax|vacc\w*|immuniz\w*)\b/i,
  },
  {
    key: "coggins",
    label: "Coggins",
    every: "yearly",
    days: 365,
    match: /\b(coggins|eia)\b/i,
  },
  {
    key: "dental",
    label: "Dental",
    every: "yearly",
    days: 365,
    match: /\b(dental|teeth|float\w*)\b/i,
  },
];

/** A live animal that could need routine care. */
export function isAnimal(h: Horse) {
  return h.status !== "off" && h.status !== "exp";
}

/** A horse the app found as a folder, before any record is attached. */
export const EMPTY_HORSE: Horse = {
  id: "", reg: null, barn: null, pending: false, idline: null, aqha: null,
  sex: null, color: null, dob: null, bought: null, boughtNote: null,
  price: null, salePrice: null, soldOn: null, status: "wip", label: null,
  pedigree: null, breeding: null, notes: [], missing: [], care: {},
  derived: {}, folderName: null, folderUrl: null, docs: [], photos: [],
  sort: 900,
};
