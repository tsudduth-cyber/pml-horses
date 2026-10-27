"use client";

import { useMemo, useState, useCallback, useEffect, useRef } from "react";
import { useRouter } from "next/navigation";
import type { Horse, CareKey, Doc, DerivedCare } from "@/lib/types";
import { STATUS_LABEL, CARE_TASKS, isAnimal } from "@/lib/types";
import {
  PageIcon,
  FolderIcon,
  ImgIcon,
  EggIcon,
  ListTab,
  CareTab,
  BreedTab,
  RecentTab,
} from "./Icons";

const SCLASS: Record<string, string> = {
  wip: "s-wip",
  done: "s-done",
  recip: "s-recip",
  exp: "s-exp",
  noreg: "s-noreg",
  off: "s-off",
};

/* ---------- small helpers ---------- */

const money = (n: number) =>
  "$" + n.toLocaleString("en-US", { maximumFractionDigits: 0 });

/**
 * The care date in force for a horse: what someone entered by hand, else
 * what the filed paperwork says.
 */
function careDate(h: Horse, k: CareKey) {
  return h.care[k] ?? h.derived[k]?.date;
}

/** Today in the viewer's own timezone, as YYYY-MM-DD. */
function today() {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function fmt(d?: string | null) {
  if (!d) return null;
  return new Date(d + "T12:00").toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
  });
}

function age(dob?: string | null) {
  if (!dob) return null;
  const d = new Date(dob + "T12:00");
  const n = new Date();
  let y = n.getFullYear() - d.getFullYear();
  if (
    n.getMonth() < d.getMonth() ||
    (n.getMonth() === d.getMonth() && n.getDate() < d.getDate())
  )
    y--;
  return y < 1 ? "Weanling" : `${y} ${y === 1 ? "yr" : "yrs"}`;
}

function daysUntil(iso: string) {
  const ms = new Date(iso + "T12:00").getTime() - Date.now();
  return Math.round(ms / 86_400_000);
}

/** When the next one is due, and how urgent that is. */
function dueFrom(last: string | undefined, intervalDays: number) {
  if (!last) return null;
  const due = new Date(last + "T12:00");
  due.setDate(due.getDate() + intervalDays);
  const iso = due.toISOString().slice(0, 10);
  const d = daysUntil(iso);
  return {
    iso,
    days: d,
    level: d < 0 ? "over" : d <= 21 ? "soon" : "ok",
    text: d < 0 ? `${-d}d over` : d === 0 ? "today" : `${d}d`,
  } as const;
}

/**
 * Photos come through the app, not straight from SharePoint. A SharePoint
 * webUrl in an <img> only renders for someone whose browser already holds a
 * SharePoint session — true on a work desktop, false on a phone, where every
 * photo would show as a broken image.
 */
function photoSrc(d: Doc, small = false) {
  if (!d.id) return d.url; // nothing better to offer
  return `/api/photo?id=${encodeURIComponent(d.id)}${small ? "&size=sm" : ""}`;
}

/**
 * A horse with no name yet is identified by her pedigree or her dam, not by
 * the word "unnamed" — that reads like it IS the name.
 */
function displayName(h: Horse) {
  if (h.pending) return h.idline || "Name pending";
  return h.barn || h.reg || h.idline || "Name pending";
}

/**
 * Which horse is already showing under each folder, so the record editor
 * can say that picking it merges rather than moves.
 */
function heldBy(horses: Horse[], exceptId: string) {
  const out: Record<string, string> = {};
  for (const h of horses) {
    if (h.id === exceptId || !h.folderName) continue;
    out[h.folderName.toLowerCase()] = displayName(h);
  }
  return out;
}

function subLine(h: Horse) {
  if (h.pending) return h.idline ?? "";
  if (h.barn && h.reg && h.barn !== h.reg) return h.reg;
  return h.pedigree ?? "";
}

function boughtText(h: Horse) {
  const d = fmt(h.bought);
  if (d && h.boughtNote) return `${d} (${h.boughtNote})`;
  return d ?? h.boughtNote ?? null;
}

function profileText(h: Horse) {
  const L: string[] = [];
  L.push(h.pending ? `Name pending — ${h.idline ?? ""}`.trim() : h.barn || h.reg || "");
  if (!h.pending && h.barn && h.reg && h.barn !== h.reg) L.push(h.reg);
  if (h.aqha) L.push("AQHA " + h.aqha);
  const sx = [h.sex, h.color].filter(Boolean).join(", ");
  if (sx) L.push(sx);
  if (h.dob) L.push(`Foaled ${fmt(h.dob)} (${age(h.dob)})`);
  if (h.pedigree) L.push(h.pedigree);
  if (h.breeding) L.push(h.breeding);
  if (h.care.coggins) L.push(`Coggins ${fmt(h.care.coggins)}`);
  return L.filter(Boolean).join("\n");
}

/* ---------- pieces ---------- */

function Mark({ h }: { h: Horse }) {
  return h.status === "exp" ? <EggIcon /> : <ImgIcon />;
}

function NameSlot({ h, big }: { h: Horse; big?: boolean }) {
  if (h.pending) return <span className="pending">NAME PENDING</span>;
  return big ? (
    <h2>{h.barn || h.reg}</h2>
  ) : (
    <div className="name">{h.barn || h.reg}</div>
  );
}

function Flag({ h }: { h: Horse }) {
  return (
    <div className={`flag ${SCLASS[h.status]}`}>
      <i className="dot" />
      {h.label || STATUS_LABEL[h.status]}
    </div>
  );
}

function DocLink({ d }: { d: Doc }) {
  return (
    <a
      className={`doc${d.folder ? " fold" : ""}`}
      href={d.url}
      target="_blank"
      rel="noopener noreferrer"
    >
      {d.folder ? <FolderIcon /> : <PageIcon />}
      <span className="dn">
        {d.name}
        {/* When a file is replaced its name does not change, so without a
            date an updated document looks untouched. */}
        {!d.folder && d.modified && (
          <span className="dm">{fmt(d.modified.slice(0, 10))}</span>
        )}
      </span>
    </a>
  );
}



/* ---------- the record, editable ---------- */

const RECORD_FIELDS = [
  { key: "barn", label: "Barn name", type: "text" },
  { key: "reg", label: "Registered name", type: "text" },
  { key: "aqha", label: "AQHA number", type: "text" },
  { key: "sex", label: "Sex", type: "sex" },
  { key: "color", label: "Color", type: "text" },
  { key: "dob", label: "Foaled", type: "date" },
  { key: "bought", label: "Bought", type: "date" },
  { key: "folderName", label: "SharePoint folder", type: "folder" },
] as const;

type RecordKey = (typeof RECORD_FIELDS)[number]["key"];

/**
 * The record, editable in place.
 *
 * Every one of these used to be seeded once and then frozen. A horse could
 * be renamed, given a folder and photographed and the app would go on
 * showing the old record, because there was no way to reach it.
 */
function RecordEditor({
  horse,
  folders,
  heldBy = {},
  onSave,
}: {
  horse: Horse;
  /** Every horse folder under /Horses/. */
  folders: string[];
  /** Folder name, lowercased, to the horse already showing under it. */
  heldBy?: Record<string, string>;
  onSave: (patch: Partial<Record<RecordKey, string | null>>) => Promise<void>;
}) {
  const current = (k: RecordKey) => (horse as any)[k] ?? "";
  const [draft, setDraft] = useState<Record<string, string>>({});
  const [busy, setBusy] = useState(false);

  const set = (k: string, v: string) => setDraft((d) => ({ ...d, [k]: v }));

  const save = async () => {
    const patch: Record<string, string | null> = {};
    for (const f of RECORD_FIELDS) {
      const was = String(current(f.key));
      const now = draft[f.key] ?? was;
      if (now !== was) patch[f.key] = now || null;
    }
    if (!Object.keys(patch).length) return;
    setBusy(true);
    await onSave(patch);
    setBusy(false);
  };

  const options = [
    ...folders,
    ...(horse.folderName && !folders.includes(horse.folderName)
      ? [horse.folderName]
      : []),
  ].sort();

  // Picking a folder another horse is already showing under folds this
  // record into that one. Say so before the tap, not after.
  const picked = draft.folderName ?? current("folderName");
  const mergeWith =
    picked && picked !== horse.folderName
      ? heldBy[String(picked).toLowerCase()]
      : undefined;

  return (
    <div className="recedit">
      {RECORD_FIELDS.map((f) => (
        <label className="rf" key={f.key}>
          <span>{f.label}</span>
          {f.type === "sex" ? (
            <select
              value={draft[f.key] ?? current(f.key)}
              disabled={busy}
              onChange={(e) => set(f.key, e.target.value)}
            >
              <option value="">—</option>
              {["Stallion", "Mare", "Gelding", "Filly", "Colt"].map((o) => (
                <option key={o}>{o}</option>
              ))}
            </select>
          ) : f.type === "folder" ? (
            <select
              value={draft[f.key] ?? current(f.key)}
              disabled={busy}
              onChange={(e) => set(f.key, e.target.value)}
            >
              <option value="">No folder</option>
              {options.map((o) => (
                <option key={o}>{o}</option>
              ))}
            </select>
          ) : (
            <input
              type={f.type === "date" ? "date" : "text"}
              value={draft[f.key] ?? current(f.key)}
              disabled={busy}
              onChange={(e) => set(f.key, e.target.value)}
            />
          )}
        </label>
      ))}
      {mergeWith && (
        <p className="rmerge">
          Saving folds this record into {mergeWith}. Notes from both are kept.
        </p>
      )}
      <div className="lacts">
        <button className="prim" disabled={busy} onClick={save}>
          Save
        </button>
      </div>
    </div>
  );
}

/* ---------- editable notes ---------- */

/**
 * Notes and the missing-items list, editable in place.
 *
 * These lived in a Word document because the app gave the office nowhere
 * to write them, and the two drifted apart. One line per note; an empty
 * line is dropped on save.
 */
function Lines({
  horse,
  field,
  label,
  lines,
  tone,
  onSave,
}: {
  horse: Horse;
  field: "notes" | "missing";
  label: string;
  lines: string[];
  tone?: "alert";
  onSave: (field: "notes" | "missing", lines: string[]) => Promise<void>;
}) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState("");
  const [busy, setBusy] = useState(false);

  const start = () => {
    setDraft(lines.join("\n"));
    setEditing(true);
  };

  const save = async () => {
    setBusy(true);
    await onSave(field, draft.split("\n"));
    setBusy(false);
    setEditing(false);
  };

  if (!editing && !lines.length) {
    return (
      <button className="addnote" onClick={start}>
        + Add {label.toLowerCase()}
      </button>
    );
  }

  return (
    <>
      <div className="lhead">
        <h3>{label}</h3>
        {!editing && (
          <button className="edit" onClick={start}>
            Edit
          </button>
        )}
      </div>

      {editing ? (
        <>
          <textarea
            className="lines"
            value={draft}
            disabled={busy}
            rows={Math.max(4, draft.split("\n").length + 1)}
            onChange={(e) => setDraft(e.target.value)}
            aria-label={`${label} for ${displayName(horse)}`}
          />
          <div className="lacts">
            <button className="prim" disabled={busy} onClick={save}>
              Save
            </button>
            <button
              className="cancel"
              disabled={busy}
              onClick={() => setEditing(false)}
            >
              Cancel
            </button>
          </div>
        </>
      ) : tone === "alert" ? (
        <ul>
          {lines.map((m, i) => (
            <li key={i}>{m}</li>
          ))}
        </ul>
      ) : (
        lines.map((n, i) => <p key={i}>{n}</p>)
      )}
    </>
  );
}

/* ---------- care editor ---------- */

function CareRow({
  horse,
  task,
  value,
  from,
  onSave,
}: {
  horse: Horse;
  task: (typeof CARE_TASKS)[number];
  value?: string;
  /** Set when the date was read off filed paperwork rather than entered. */
  from?: DerivedCare;
  onSave: (key: CareKey, date: string | null) => Promise<void>;
}) {
  const [busy, setBusy] = useState(false);
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState("");
  const due = dueFrom(value, task.days);

  const commit = async (date: string | null) => {
    setBusy(true);
    await onSave(task.key, date);
    setBusy(false);
    setOpen(false);
  };

  return (
    <div className="carerow">
      <button
        className="caredate"
        aria-expanded={open}
        disabled={busy}
        onClick={() => {
          setDraft(value ?? today());
          setOpen((v) => !v);
        }}
      >
        <span className="l">{task.label}</span>
        <span className="r">
          {due && (
            <span className={`due ${due.level}`}>
              {due.level === "over" ? "OVERDUE " : "due "}
              {due.text}
            </span>
          )}
          <span className={value ? (from ? "val derived" : "val") : "val empty"}>
            {value ? fmt(value) : "not on file"}
          </span>
        </span>
      </button>

      {from && !open && (
        <div className="fromfile">
          from {from.from}
          {from.barn && " (whole barn)"}
        </div>
      )}

      {open && (
        <div className="careedit">
          <div className="hint">
            {from
              ? `Read from ${from.from}${from.barn ? " (whole barn)" : ""}. Entering a date here replaces it.`
              : `${task.label} last done — ${task.every}`}
          </div>
          <button className="prim" disabled={busy} onClick={() => commit(today())}>
            Done today
          </button>
          <div className="pick">
            <input
              type="date"
              aria-label={`${task.label} date, ${displayName(horse)}`}
              value={draft}
              max={today()}
              disabled={busy}
              onChange={(e) => setDraft(e.target.value)}
            />
            <button
              disabled={busy || !draft}
              onClick={() => commit(draft)}
            >
              Save
            </button>
          </div>
          <div className="acts">
            {/* Only a hand-entered date can be cleared. Clearing a derived
                one would mean deleting the paperwork, which is not this
                app's job — it falls back to the file on its own. */}
            {horse.care[task.key] && (
              <button className="clear" disabled={busy} onClick={() => commit(null)}>
                {from ? "Use the filed date instead" : "Clear date"}
              </button>
            )}
            <button className="cancel" disabled={busy} onClick={() => setOpen(false)}>
              Cancel
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------- share sheet ---------- */

function ShareSheet({
  horse,
  onClose,
  toast,
}: {
  horse: Horse;
  onClose: () => void;
  toast: (m: string) => void;
}) {
  const body = profileText(horse);
  const name = displayName(horse);
  const enc = encodeURIComponent(body);

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(body);
      toast("Copied");
    } catch {
      toast("Select the text and copy");
    }
  };

  return (
    <>
      <div className="scrim" onClick={onClose} />
      <div className="sheet" role="dialog" aria-modal="true">
        <div className="sh">
          <h4>Send profile</h4>
          <button className="x" onClick={onClose} aria-label="Close">
            ×
          </button>
        </div>
        <pre>{body}</pre>
        <div className="go">
          <button className="btn" onClick={copy}>
            Copy
          </button>
          <a className="btn ghost" href={`sms:?&body=${enc}`}>
            Text
          </a>
          <a
            className="btn ghost"
            href={`mailto:?subject=${encodeURIComponent(name)}&body=${enc}`}
          >
            Email
          </a>
        </div>
        <div className="note">Price is left out on purpose.</div>
      </div>
    </>
  );
}

/* ---------- detail ---------- */

function Detail({
  horse,
  showMoney,
  onBack,
  onSave,
  onSaveText,
  onSaveRecord,
  folders,
  heldBy,
  onShare,
}: {
  horse: Horse;
  showMoney: boolean;
  onBack: () => void;
  onSave: (key: CareKey, date: string | null) => Promise<void>;
  onSaveText: (field: "notes" | "missing", lines: string[]) => Promise<void>;
  onSaveRecord: (patch: Record<string, string | null>) => Promise<void>;
  folders: string[];
  heldBy: Record<string, string>;
  onShare: () => void;
}) {
  const [shot, setShot] = useState(0);
  const [editRec, setEditRec] = useState(false);
  const h = horse;

  const rows: [string, string | null, string][] = [
    ["Registered name", h.reg, "Not named yet"],
    ["AQHA number", h.aqha, "Missing"],
    ["Sex", h.sex, "Not recorded"],
  ];
  const ped = h.pending ? h.idline : h.pedigree;
  if (ped) rows.splice(1, 0, ["Pedigree", ped, "—"]);
  if (h.color) rows.push(["Color", h.color, "—"]);
  rows.push(
    ["Foaled", fmt(h.dob), "Missing"],
    ["Age", age(h.dob), "—"],
    ["Bought", boughtText(h), "Missing"]
  );
  if (h.status === "off") rows.push(["Sold", fmt(h.soldOn), "Missing"]);
  if (showMoney) {
    rows.push([
      "Purchase price",
      h.price != null ? money(h.price) : null,
      "Missing",
    ]);
    if (h.status === "off")
      rows.push([
        "Sale price",
        h.salePrice != null ? money(h.salePrice) : null,
        "Missing",
      ]);
  }

  const hero = h.photos[shot];

  return (
    <>
      <button className="back" onClick={onBack}>
        ← All horses
      </button>

      <div className="photo">
        {hero ? (
          // eslint-disable-next-line @next/next/no-img-element
          <img src={photoSrc(hero)} alt={displayName(h)} />
        ) : (
          <>
            <Mark h={h} />
            <span>NO PHOTO ON FILE</span>
          </>
        )}
      </div>

      {h.photos.length > 1 && (
        <div className="shots">
          {h.photos.map((p, i) => (
            <button
              key={p.url}
              aria-current={i === shot}
              onClick={() => setShot(i)}
            >
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img src={photoSrc(p, true)} alt={p.name} />
            </button>
          ))}
        </div>
      )}

      <div className="dhead">
        <NameSlot h={h} big />
        {h.pending ? (
          <div className="reg" style={{ marginTop: 7 }}>
            {h.idline}
          </div>
        ) : (
          subLine(h) && <div className="reg">{subLine(h)}</div>
        )}
        <div style={{ marginTop: 10 }}>
          <Flag h={h} />
        </div>
      </div>

      <section className="sec">
        <div className="lhead">
          <h3>RECORD</h3>
          <button className="edit" onClick={() => setEditRec((v) => !v)}>
            {editRec ? "Done" : "Edit"}
          </button>
        </div>
        {editRec ? (
          <RecordEditor
            horse={h}
            folders={folders}
            heldBy={heldBy}
            onSave={async (p) => {
              await onSaveRecord(p as Record<string, string | null>);
              setEditRec(false);
            }}
          />
        ) : (
          <dl>
            {rows.map(([k, v, fallback]) => (
              <div className="kv" key={k}>
                <dt>{k}</dt>
                <dd className={v ? "" : "miss"}>{v || fallback}</dd>
              </div>
            ))}
          </dl>
        )}
      </section>

      {h.breeding && (
        <section className="sec breed">
          <h3>BREEDING</h3>
          <div className="bline">{h.breeding}</div>
        </section>
      )}

      {isAnimal(h) && (
        <section className="sec">
          <h3>ROUTINE CARE</h3>
          {CARE_TASKS.map((t) => (
            <CareRow
              key={t.key}
              horse={h}
              task={t}
              value={h.care[t.key] ?? h.derived[t.key]?.date}
              from={h.care[t.key] ? undefined : h.derived[t.key]}
              onSave={onSave}
            />
          ))}
        </section>
      )}

      <section className="sec notes">
        <Lines
          horse={h}
          field="notes"
          label="NOTES"
          lines={h.notes}
          onSave={onSaveText}
        />
      </section>

      <section className="sec">
        <div className={h.missing.length ? "missing" : undefined}>
          <Lines
            horse={h}
            field="missing"
            label="STILL MISSING"
            lines={h.missing}
            tone="alert"
            onSave={onSaveText}
          />
        </div>
      </section>

      {h.docs.length > 0 && (
        <section className="sec">
          <h3>DOCUMENTS</h3>
          {h.docs.map((d) => (
            <DocLink key={d.url} d={d} />
          ))}
        </section>
      )}

      <div className="actions">
        <button className="btn" onClick={onShare}>
          Send profile
        </button>
        {h.folderUrl ? (
          <a
            className="btn ghost"
            href={h.folderUrl}
            target="_blank"
            rel="noopener noreferrer"
          >
            Open folder
          </a>
        ) : (
          <button className="btn ghost" disabled>
            No folder yet
          </button>
        )}
      </div>
    </>
  );
}

/* ---------- the app ---------- */

const FILTERS: [string, string, (h: Horse) => boolean][] = [
  ["all", "All", (h) => h.status !== "off"],
  ["wip", "In progress", (h) => h.status === "wip"],
  ["reg", "Registered", (h) => h.status === "done"],
  [
    "breed",
    "Breeding",
    (h) => !!h.breeding || h.status === "exp" || h.status === "recip",
  ],
  ["off", "Sold", (h) => h.status === "off"],
];

type View = "horses" | "detail" | "care" | "breeding" | "recent";

export default function App({
  horses: initial,
  season,
  rescan,
  folders = [],
}: {
  horses: Horse[];
  season: { when: string; what: string; detail: string; hot: boolean }[];
  /** The saved reading of the paperwork has gone stale. */
  rescan?: boolean;
  /** Every horse folder under /Horses/, for the folder picker. */
  folders?: string[];
}) {
  const [horses, setHorses] = useState(initial);
  const router = useRouter();

  // Fresh data from the server replaces what is on screen. useState only
  // reads its argument on first render, so without this a refresh fetched
  // new horses and the page went on showing the old ones.
  useEffect(() => setHorses(initial), [initial]);

  /**
   * Re-fetch when the app comes back to the front.
   *
   * On iOS a home-screen app resuming from the background restores the last
   * rendered screen rather than reloading, so Philip could tap the icon and
   * read days-old data with nothing telling him it was stale. Refresh on
   * return, and at most once a minute so flicking between apps does not
   * hammer SharePoint.
   */
  const lastPull = useRef(Date.now());
  useEffect(() => {
    const pull = () => {
      if (document.visibilityState !== "visible") return;
      if (Date.now() - lastPull.current < 60_000) return;
      lastPull.current = Date.now();
      router.refresh();
    };
    document.addEventListener("visibilitychange", pull);
    window.addEventListener("focus", pull);
    window.addEventListener("pageshow", pull);
    return () => {
      document.removeEventListener("visibilitychange", pull);
      window.removeEventListener("focus", pull);
      window.removeEventListener("pageshow", pull);
    };
  }, [router]);

  /**
   * Re-read the paperwork when the saved reading is old.
   *
   * It runs here rather than on a schedule because SharePoint access is in
   * the name of the person signed in, and a scheduled job has nobody to be.
   * Whoever opens the app pays for it, in the background, and the page
   * refreshes once it lands.
   */
  const scanning = useRef(false);
  useEffect(() => {
    if (!rescan || scanning.current) return;
    scanning.current = true;
    fetch("/api/scan", { method: "POST" })
      .then((r) => (r.ok ? router.refresh() : null))
      .catch(() => {});
  }, [rescan, router]);

  const [view, setView] = useState<View>("horses");
  const [filter, setFilter] = useState("all");
  const [openId, setOpenId] = useState<string | null>(null);
  const [showMoney, setShowMoney] = useState(false);
  const [shareId, setShareId] = useState<string | null>(null);
  const [msg, setMsg] = useState<string | null>(null);

  const toast = useCallback((m: string) => {
    setMsg(m);
    setTimeout(() => setMsg(null), 1800);
  }, []);

  const open = horses.find((h) => h.id === openId) ?? null;
  const sharing = horses.find((h) => h.id === shareId) ?? null;
  const herd = useMemo(() => horses.filter(isAnimal), [horses]);

  /* save a care date — optimistic, reverted if SharePoint refuses */
  const saveCare = useCallback(
    async (horseId: string, key: CareKey, date: string | null) => {
      const before = horses.find((h) => h.id === horseId)?.care[key];
      setHorses((hs) =>
        hs.map((h) => {
          if (h.id !== horseId) return h;
          const care = { ...h.care };
          if (date) care[key] = date;
          else delete care[key];
          return { ...h, care };
        })
      );
      try {
        const res = await fetch("/api/care", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: horseId, field: key, date }),
        });
        if (!res.ok) throw new Error((await res.json()).error ?? "Failed");
        toast("Saved");
      } catch (e: any) {
        setHorses((hs) =>
          hs.map((h) => {
            if (h.id !== horseId) return h;
            const care = { ...h.care };
            if (before) care[key] = before;
            else delete care[key];
            return { ...h, care };
          })
        );
        toast(e?.message ?? "Didn't save");
      }
    },
    [horses, toast]
  );

  /* Notes and missing items — optimistic, reverted if SharePoint refuses */
  const saveText = useCallback(
    async (horseId: string, field: "notes" | "missing", lines: string[]) => {
      const before = horses.find((h) => h.id === horseId)?.[field] ?? [];
      const clean = lines.map((l) => l.trim()).filter(Boolean);
      setHorses((hs) =>
        hs.map((h) => (h.id === horseId ? { ...h, [field]: clean } : h))
      );
      try {
        const res = await fetch("/api/notes", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: horseId, field, lines: clean }),
        });
        if (!res.ok) throw new Error((await res.json()).error ?? "Failed");
        toast("Saved");
      } catch (e: any) {
        setHorses((hs) =>
          hs.map((h) => (h.id === horseId ? { ...h, [field]: before } : h))
        );
        toast(e?.message ?? "Didn't save");
      }
    },
    [horses, toast]
  );

  /* Record fields — a reload follows, because a rename changes the list */
  const saveRecord = useCallback(
    async (horseId: string, patch: Record<string, string | null>) => {
      try {
        const res = await fetch("/api/record", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: horseId, fields: patch }),
        });
        if (!res.ok) throw new Error((await res.json()).error ?? "Failed");
        toast("Saved");
        router.refresh();
      } catch (e: any) {
        toast(e?.message ?? "Didn't save");
      }
    },
    [toast, router]
  );

  /* Recent — built from the real file dates in SharePoint */
  const recent = useMemo(() => {
    const all: { date: string; name: string; horse: string; url: string }[] = [];
    for (const h of horses) {
      for (const d of [...h.docs, ...h.photos]) {
        if (!d.modified) continue;
        all.push({
          date: d.modified.slice(0, 10),
          name: d.name,
          horse: displayName(h),
          url: d.url,
        });
      }
    }
    all.sort((a, b) => b.date.localeCompare(a.date));
    const days: { day: string; items: typeof all }[] = [];
    for (const it of all.slice(0, 40)) {
      const label = fmt(it.date) ?? it.date;
      const last = days[days.length - 1];
      if (last && last.day === label) last.items.push(it);
      else days.push({ day: label, items: [it] });
    }
    return days;
  }, [horses]);

  /* Care — what's overdue or coming, across the herd */
  const upcoming = useMemo(() => {
    const out: {
      horse: Horse;
      task: (typeof CARE_TASKS)[number];
      due: NonNullable<ReturnType<typeof dueFrom>>;
    }[] = [];
    for (const h of herd) {
      for (const t of CARE_TASKS) {
        const d = dueFrom(careDate(h, t.key), t.days);
        if (d && d.days <= 45) out.push({ horse: h, task: t, due: d });
      }
    }
    return out.sort((a, b) => a.due.days - b.due.days);
  }, [herd]);

  const breeders = useMemo(
    () => horses.filter((h) => h.breeding),
    [horses]
  );

  const goDetail = (id: string) => {
    setOpenId(id);
    setView("detail");
    window.scrollTo(0, 0);
  };

  return (
    <div className="shell">
      <header className="mast">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src="/pml-logo.png" alt="PML Performance Horses" />
        <button
          className="money"
          aria-pressed={showMoney}
          onClick={() => setShowMoney((v) => !v)}
        >
          {showMoney ? "Hide prices" : "Show prices"}
        </button>
      </header>

      {/* HORSES */}
      <main className={`view${view === "horses" ? "" : " hidden"}`}>
        <div className="filters">
          {FILTERS.map(([key, label, test]) => (
            <button
              key={key}
              className="chip"
              aria-pressed={filter === key}
              onClick={() => setFilter(key)}
            >
              {label}
              <b>{horses.filter(test).length}</b>
            </button>
          ))}
        </div>
        <div className="list">
          {horses
            .filter(FILTERS.find((f) => f[0] === filter)![2])
            .map((h) => {
              const bits = [
                h.sex,
                h.aqha ? "AQHA " + h.aqha : null,
                age(h.dob),
              ]
                .filter(Boolean)
                .join(" · ");
              const sub = subLine(h);
              return (
                <button
                  key={h.id}
                  className="row"
                  onClick={() => goDetail(h.id)}
                >
                  <div className="thumb">
                    {h.photos[0] ? (
                      // eslint-disable-next-line @next/next/no-img-element
                      <img src={photoSrc(h.photos[0], true)} alt="" />
                    ) : (
                      <Mark h={h} />
                    )}
                  </div>
                  <div className="rowbody">
                    <NameSlot h={h} />
                    {sub && <div className="reg">{sub}</div>}
                    {bits && <div className="meta">{bits}</div>}
                    <Flag h={h} />
                    {h.missing.length > 0 && (
                      <div className="gapline">
                        {h.missing.length} thing
                        {h.missing.length > 1 ? "s" : ""} missing
                      </div>
                    )}
                  </div>
                  {showMoney &&
                    (h.status === "off" ? (
                      h.salePrice != null ? (
                        <div className="price">
                          {money(h.salePrice)}
                          <small>SOLD</small>
                        </div>
                      ) : (
                        <div className="price na">sale price?</div>
                      )
                    ) : h.price != null ? (
                      <div className="price">{money(h.price)}</div>
                    ) : (
                      <div className="price na">—</div>
                    ))}
                </button>
              );
            })}
        </div>
      </main>

      {/* DETAIL */}
      <main className={`view detail${view === "detail" ? "" : " hidden"}`}>
        {open && (
          <Detail
            key={open.id}
            horse={open}
            showMoney={showMoney}
            onBack={() => setView("horses")}
            onSave={(k, d) => saveCare(open.id, k, d)}
            onSaveText={(f, l) => saveText(open.id, f, l)}
            onSaveRecord={(p) => saveRecord(open.id, p)}
            folders={folders}
            heldBy={heldBy(horses, open.id)}
            onShare={() => setShareId(open.id)}
          />
        )}
      </main>

      {/* CARE */}
      <main className={`view care${view === "care" ? "" : " hidden"}`}>
        {upcoming.length > 0 &&
          upcoming.map(({ horse, task, due }) => (
            <div
              className={`when${due.level === "over" ? " hot" : ""}`}
              key={horse.id + task.key}
            >
              <div className="d">
                {due.level === "over" ? "Overdue" : fmt(due.iso)}
              </div>
              <div className="b">
                <div className="h">
                  {task.label} — {displayName(horse)}
                </div>
                <div className="s">
                  Last done {fmt(horse.care[task.key])}. {task.every}.
                </div>
              </div>
            </div>
          ))}

        {CARE_TASKS.map((t) => {
          const have = herd.filter((h) => careDate(h, t.key));
          const pct = herd.length
            ? Math.round((have.length / herd.length) * 100)
            : 0;
          return (
            <div className="task" key={t.key}>
              <div className="top">
                <div className="t">{t.label}</div>
                <div className="iv">{t.every}</div>
              </div>
              <div className="bar">
                <i style={{ width: `${pct}%` }} />
              </div>
              <div className="sum">
                <b>
                  {have.length} of {herd.length}
                </b>{" "}
                have a date on file.
              </div>
              {have.length > 0 && (
                <div className="who">
                  {have.map((h) => {
                    const d = dueFrom(careDate(h, t.key), t.days)!;
                    return (
                      <button
                        key={h.id}
                        className={d.level}
                        onClick={() => goDetail(h.id)}
                      >
                        {displayName(h)}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </main>

      {/* BREEDING */}
      <main className={`view breeding${view === "breeding" ? "" : " hidden"}`}>
        {breeders.map((h) => (
          <div className="when" key={h.id}>
            <div className="d">{displayName(h)}</div>
            <div className="b">
              <div className="h">{h.breeding}</div>
              {h.care.coggins ? (
                <div className="s">Coggins {fmt(h.care.coggins)}.</div>
              ) : (
                <div className="s" style={{ color: "var(--alert)" }}>
                  No Coggins on file — needed before she ships.
                </div>
              )}
            </div>
          </div>
        ))}
        <section className="sec">
          <h3>SEASON</h3>
          {season.map((s) => (
            <div className="kv" key={s.when}>
              <dt>{s.when}</dt>
              <dd className={s.hot ? "miss" : ""}>{s.what}</dd>
            </div>
          ))}
        </section>
      </main>

      {/* RECENT */}
      <main className={`view recent${view === "recent" ? "" : " hidden"}`}>
        {recent.map((d) => (
          <div key={d.day}>
            <div className="day">{d.day.toUpperCase()}</div>
            {d.items.map((it) => (
              <a
                className="ev"
                key={it.url}
                href={it.url}
                target="_blank"
                rel="noopener noreferrer"
              >
                <PageIcon />
                <div>
                  <div className="w">{it.name}</div>
                  <div className="h">{it.horse}</div>
                </div>
              </a>
            ))}
          </div>
        ))}
      </main>

      <nav className="tabs" role="tablist">
        {(
          [
            ["horses", "HORSES", <ListTab key="a" />],
            ["care", "CARE", <CareTab key="b" />],
            ["breeding", "BREEDING", <BreedTab key="c" />],
            ["recent", "RECENT", <RecentTab key="d" />],
          ] as const
        ).map(([key, label, icon]) => (
          <button
            key={key}
            className="tab"
            role="tab"
            aria-selected={
              view === key || (view === "detail" && key === "horses")
            }
            onClick={() => {
              setView(key as View);
              window.scrollTo(0, 0);
            }}
          >
            {icon}
            {label}
          </button>
        ))}
      </nav>

      {sharing && (
        <ShareSheet
          horse={sharing}
          onClose={() => setShareId(null)}
          toast={toast}
        />
      )}
      {msg && <div className="toast">{msg}</div>}
    </div>
  );
}
