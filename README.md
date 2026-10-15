# PML Performance Horses

Philip's horse records on his phone. Reads SharePoint, writes care dates back.

Next.js 15 on Vercel. Entra ID sign-in, Microsoft Graph for data. No separate
database — SharePoint is the system of record.

---

## How it works

**Documents** come from the horse folders under
`Shared Documents / Company Data / Lang Interests LLC / General / _Owned Properties / PML Performance Horses LLC / Horses`.
The app lists what's there and links straight into SharePoint. A `Photos`
subfolder becomes the horse's pictures. Nothing is copied.

**Facts** — AQHA number, sex, DOB, prices, status, notes, care dates — come
from a SharePoint list called **PML Horses**. That list replaces
`Horse Info & Registration Status.docx`.

**Writes** go one direction only. Philip can set the five care dates from the
app. Everything else is read-only, so the filing and the paperwork stay with
whoever keeps them.

Each horse's row points at its folder by name (`FolderName`), so renaming a
folder only means updating that one cell.

---

## Setup

Four steps. Step 1 needs an Azure admin. Do step 3 before step 2 — the app
builds the list for you, so it has to be running first.

### 1. Register the app in Entra ID

Azure portal → Microsoft Entra ID → App registrations → New registration.

- Name: `PML Horses`
- Accounts: this organizational directory only
- Redirect URI: Web → `https://YOUR-APP.vercel.app/api/auth/callback/azure-ad`

Add a second redirect URI for local work:
`http://localhost:3000/api/auth/callback/azure-ad`

Then:

- **Certificates & secrets** → new client secret → copy the value now, it is
  shown once.
- **API permissions** → Microsoft Graph → Delegated → `Sites.ReadWrite.All`,
  plus `openid`, `profile`, `email`, `offline_access`. Grant admin consent.

Copy the Application (client) ID and the Directory (tenant) ID.

> Delegated, not application, permissions. Philip signs in as himself and the
> app can only see what he can already see.

### 2. Create the SharePoint list

The app builds it for you. Finish step 3 first, start the app, sign in, then
open:

```
http://localhost:3000/api/setup
```

That page tells you what it is about to do and changes nothing. Add
`?confirm=yes` and it creates the `PML Horses` list with all 26 columns and
loads all 26 horses.

It is safe to run twice — it will not create a second list, and it will not
write over rows that are already there.

If you would rather build it by hand, the columns are below and
`setup/pml-horses-seed.csv` has the data. Open the CSV in Excel, copy the 26
data rows, and paste into the list's **Edit in grid view**.

| Column | Type | Notes |
| --- | --- | --- |
| `Title` | Single line | Registered name. Blank if not registered yet. |
| `BarnName` | Single line | |
| `NamePending` | Yes/No | Yes = app shows "NAME PENDING", not a name |
| `IdLine` | Single line | How to identify a pending horse — pedigree, or who she's out of |
| `AQHA` | Single line | |
| `Sex` | Choice | Stallion, Mare, Gelding, Filly, Colt |
| `Color` | Single line | |
| `Foaled` | Date | Date only |
| `Bought` | Date | Date only |
| `BoughtNote` | Single line | e.g. "unconfirmed", "month only" |
| `Price` | Currency | Purchase price |
| `SalePrice` | Currency | What it sold for |
| `SoldOn` | Date | |
| `Status` | Choice | Work in progress, Registered to PML, Recip mare, Embryo, Not registering, Sold |
| `StatusNote` | Single line | Overrides the status label when it needs to read differently |
| `Pedigree` | Single line | |
| `BreedingTo` | Single line | e.g. "Bred to LBR Hangem (Shotty) at Solo Select" |
| `Notes` | Multiple lines, plain text | One note per line |
| `Missing` | Multiple lines, plain text | One item per line |
| `Farrier` | Date | Last done |
| `Deworm` | Date | Last done |
| `Vax` | Date | Last done |
| `Coggins` | Date | Last done |
| `Dental` | Date | Last done |
| `FolderName` | Single line | Folder name under `/Horses/`. Sold horses use `x_Sold Horses/Name` |
| `Sort` | Number | Display order |

Set every Date column to **Date Only**, not Date and Time.

In `Notes` and `Missing`, a ` | ` separates one item from the next. The app
renders each as its own line. It reads real line breaks too, so when somebody
types into those fields in SharePoint later they can just hit return.

### 3. Environment variables

Copy `.env.example` to `.env.local` and fill in the top two blocks. The
SharePoint block is already correct for Lang Interests.

```
AZURE_AD_CLIENT_ID=        from step 1
AZURE_AD_CLIENT_SECRET=    from step 1
AZURE_AD_TENANT_ID=        from step 1
NEXTAUTH_URL=              http://localhost:3000 locally
NEXTAUTH_SECRET=           openssl rand -base64 32
```

```bash
npm install
npm run dev
```

### 4. Deploy

Push to GitHub, import into Vercel, paste the same variables into Vercel's
environment settings. Set `NEXTAUTH_URL` to the real Vercel URL, and make sure
that URL's callback is in the Entra redirect list.

On Philip's phone: open the site in Safari, Share → Add to Home Screen. It
installs as an app with the PML icon and opens without the address bar.

---

## What Philip can change

Only the five care dates, by tapping the date field on a horse. Those write
straight back to the list. Everything else is read-only in the app and gets
changed in SharePoint.

A save that fails rolls the date back on screen and says so, rather than
looking saved and quietly not being.

---

## Notes

**Care intervals** are in `src/lib/types.ts` — farrier every 7 weeks,
deworming quarterly, vaccinations, Coggins and dental yearly. These are
standard defaults, not Philip's. Confirm them with his vet and farrier before
anybody trusts a due date.

**The breeding season calendar** is in `src/lib/season.ts`. It covers the
operation rather than any one horse, so it lives in code. Edit it each year.

**Prices** are hidden by default and shown with the header button. That's a
courtesy so Philip can hand someone the phone, not a security control —
anyone holding it can tap the button. If it needs to be locked, that's a
different build.

**No offline queue.** If a save fails it tells you and rolls back. Adding a
real offline queue is maybe an hour of work in `App.tsx` if signal at the
barn turns out to be a problem.

**Photo order** is alphabetical by filename, and the first one becomes the
horse's picture. Name a shot `01 Profile.jpg` to pin it first.
