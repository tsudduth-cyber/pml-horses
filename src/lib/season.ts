/**
 * The breeding-season calendar, from the breeding plan in SharePoint.
 * These are dates for the operation as a whole, not per horse, so they live
 * here rather than in the list. Edit this file each year.
 */
export const SEASON: {
  when: string;
  what: string;
  detail: string;
  hot: boolean;
}[] = [
  {
    when: "Mid-Nov",
    what: "Start core vaccinations",
    detail:
      "KS recommends starting mid-November to clear the holidays and finish before the first week of February.",
    hot: true,
  },
  {
    when: "Thanksgiving",
    what: "Mares under lights",
    detail:
      "16 hours of total light a day. Only needed if breeding in early February.",
    hot: false,
  },
  {
    when: "Dec – Jan",
    what: "Embryo 2 moves to Weatherford Equine",
    detail: "From EquiEmbryo.",
    hot: false,
  },
  {
    when: "January",
    what: "Philip wants to breed",
    detail: "His stated preference, if it can be done.",
    hot: false,
  },
  {
    when: "Feb 1",
    what: "Breeding season opens",
    detail:
      "Runs through Jul 31. Weatherford Equine takes haul-in palpations 8–10am, seven days a week, no appointment.",
    hot: false,
  },
  {
    when: "Mid-March",
    what: "Fallback — natural cycling",
    detail: "If not breeding early, start cycle checks mid-March.",
    hot: false,
  },
  {
    when: "Spring 2027",
    what: "Embryo 2 transfer",
    detail: "Into the Appaloosa cross mare.",
    hot: false,
  },
];
