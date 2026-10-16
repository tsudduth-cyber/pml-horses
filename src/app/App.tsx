"use client";

import { useMemo, useState, useCallback } from "react";
import type { Horse, CareKey, Doc } from "@/lib/types";
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
 * A horse with no name yet is identified by her pedigree or her dam, not by
 * the word "unnamed" — that reads like it IS the name.
 */
function displayName(h: Horse) {
  if (h.pending) return h.idline || "Name pending";
  return h.barn || h.reg || h.idline || "Name pending";
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
      {d.name}
    </a>
  );
}

/* ---------- care editor ---------- */

function CareRow({
  horse,
  task,
  value,
  onSave,
}: {
  horse: Horse;
  task: (typeof CARE_TASKS)[number];
  value?: string;
  onSave: (key: CareKey, date: string | null) => Promise<void>;
}) {
  const [busy, setBusy] = useState(false);
  const due = dueFrom(value, task.days);

  return (
    <div className="caredate">
      <span className="l">{task.label}</span>
      <span className="r">
        {due && (
          <span className={`due ${due.level}`}>
            {due.level === "over" ? "OVERDUE " : "due "}
            {due.text}
          </span>
        )}
        <input
          type="date"
          aria-label={`${task.label} last done, ${displayName(horse)}`}
          className={value ? "" : "empty"}
          value={value ?? ""}
          disabled={busy}
          onChange={async (e) => {
            setBusy(true);
            await onSave(task.key, e.target.value || null);
            setBusy(false);
          }}
        />
      </span>
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
  onShare,
}: {
  horse: Horse;
  showMoney: boolean;
  onBack: () => void;
  onSave: (key: CareKey, date: string | null) => Promise<void>;
  onShare: () => void;
}) {
  const [shot, setShot] = useState(0);
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
          // Served straight from SharePoint; these URLs are short-lived, which
          // is fine because the page is re-fetched on every load.
          // eslint-disable-next-line @next/next/no-img-element
          <img src={hero.url} alt={displayName(h)} />
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
              <img src={p.url} alt={p.name} />
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
        <h3>RECORD</h3>
        <dl>
          {rows.map(([k, v, fallback]) => (
            <div className="kv" key={k}>
              <dt>{k}</dt>
              <dd className={v ? "" : "miss"}>{v || fallback}</dd>
            </div>
          ))}
        </dl>
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
              value={h.care[t.key]}
              onSave={onSave}
            />
          ))}
        </section>
      )}

      {h.notes.length > 0 && (
        <section className="sec notes">
          <h3>NOTES</h3>
          {h.notes.map((n, i) => (
            <p key={i}>{n}</p>
          ))}
        </section>
      )}

      {h.missing.length > 0 && (
        <section className="sec">
          <div className="missing">
            <h4>STILL MISSING</h4>
            <ul>
              {h.missing.map((m, i) => (
                <li key={i}>{m}</li>
              ))}
            </ul>
          </div>
        </section>
      )}

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
}: {
  horses: Horse[];
  season: { when: string; what: string; detail: string; hot: boolean }[];
}) {
  const [horses, setHorses] = useState(initial);
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
        const d = dueFrom(h.care[t.key], t.days);
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
        <div>
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
                      <img src={h.photos[0].url} alt="" />
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
      <main className={`view${view === "detail" ? "" : " hidden"}`}>
        {open && (
          <Detail
            key={open.id}
            horse={open}
            showMoney={showMoney}
            onBack={() => setView("horses")}
            onSave={(k, d) => saveCare(open.id, k, d)}
            onShare={() => setShareId(open.id)}
          />
        )}
      </main>

      {/* CARE */}
      <main className={`view${view === "care" ? "" : " hidden"}`}>
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
          const have = herd.filter((h) => h.care[t.key]);
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
                    const d = dueFrom(h.care[t.key], t.days)!;
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
      <main className={`view${view === "breeding" ? "" : " hidden"}`}>
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
      <main className={`view${view === "recent" ? "" : " hidden"}`}>
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
