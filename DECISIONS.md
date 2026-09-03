# Why the code is the way it is

Context a reviewer needs that does not belong in the source. Code comments state the *rule*;
this file holds the *history* — what went wrong, when, and why the fix took the shape it did.

Referenced from `app.js` as `see DECISIONS.md#<anchor>`.

---

## invoiceable-status

**Rule:** only an **Active** client is invoiced. "In Progress" (stored as `inactive`), Lost and
Terminated are excluded from every invoicing surface — generation, the missing-invoice report, the
monthly preview and the batch send.

**Why one predicate:** those four surfaces each carried their own inline status test and had
drifted. `updateMissingReport` relied on `clientDueForInvoice`, which never checked status at all,
so it nagged about clients who were not being served yet. Owner decision, 2026-08-21.

Note `inactive` is **displayed** as "In Progress" everywhere. Exports that wrote the raw value
reported mid-onboarding clients as dropped (fixed 2026-09-03).

## missing-invoice-start-date

A client is only "missing an invoice" for a period if their service start date is on or before that
period. No start date means we cannot say they were active — do not nag. Clients with no start
date, or one after the period, were being flagged.

## carrier-clients

Managed-care (carrier) clients are billed through the carrier's own software and never receive a
CRM invoice, so they are excluded from every invoicing surface. A CHAMPS client assigned to a
carrier caseworker is a separate hazard: the invoice would email Medicaid PHI to a private insurer,
so the caseworker's organisation is checked at send time too.

## authorization-gates

`hasAuthorization` is true on hours, tasks **or** an effective date — enough to make the agency the
client's provider, which is what gates Active status.

`hasBillableAuthorization` is stricter: it requires approved **hours**, because an authorization
carrying only an effective date (a partial OCR read) generated a blank-total invoice, and that
invoice then satisfied the "already has an invoice for this period" check — so the client silently
stopped showing as missing one.

## current-record-across-async

**The most important entry here.** The app names the record being worked on in a global —
`activeProfileName`, `activeCgId`, `activeCwId`. Reading it synchronously is safe. Reading it
**after an async gap** — a fetch, or a confirm dialog waiting on a click — is not: the owner can
navigate away in between, and the write lands on whoever is current *then*.

Because the DOM is reused rather than rebuilt (`cg-milogin-pass` is one static element cleared
between caregivers), a late write goes into a **live** field belonging to someone else.

This single mistake produced, among others:

| What happened | Fix |
|---|---|
| One client's documents rendered under another client's name; the extract, email and delete buttons then acted on the wrong client | #159 |
| A caregiver's MI Login password written into another caregiver's form, and saved onto their record | #179 |
| An invoice marked Paid on the wrong client, or the wrong month | #167 |
| An authorization cleared from the wrong client | #176 |
| Invoice save overwriting a different month's submitted invoice | #152 |
| Auto-generate duplicating an invoice | #153 |
| Roster deletes reverting concurrent edits | #156 |
| Assistant confirmations replacing one another | #157 |
| A deleted task returning on the next sync | #161 |

**The rule now:** capture the record id *before* the gap, then `whenStillOn(kind, id, fn)`.
`stillOn` fails closed — an unrecognised kind returns false, because defaulting to true writes to
whatever happens to be on screen. `tests/no-stale-record-writes.test.js` fails the build if a new
instance appears.

## invoice-immutability

An invoice is a certified statement, not a live view. Anything re-derived at print time can rewrite
a document that was already sent, so the invoice stores its own values and replays them:

- **Rate** — re-deriving today's state rate showed the current rate on a prior-year invoice, and
  saving wrote that rate over the original.
- **Bill To** — re-derived from the caseworker's *current* agency, so reassigning a caseworker
  rewrote the Bill To line on every historic invoice the moment it was opened. Generated invoices
  did not store one at all until 2026-09-03, so they always took the fallback.
- **Signature** — `sigId` records which signature certified the invoice. Reopening one used to
  clear the signature and not restore it, so a reprint fell back to the first signature in the
  library — a different named person — and the blank was written back over the record.

## billing-rounding

Two documents round in **opposite directions**, deliberately (owner, 2026-09-01):

- **The invoice bills the authorization exactly.** "The number should be exact on the invoice. We
  can't over bill." Billing over the authorization is what triggers a recoupment.
- **The caregiver task sheet is padded slightly over** — up to the next half hour, minimum 15
  minutes of headroom. The agency must never deliver less time than was authorized.

A test pins the invoice as unpadded. Do not "fix" them into agreement.

## authorization-totals-disagree

A DHS-1210-A packet can state the approved monthly total twice and the figures can differ. One real
packet reads "62 Hours and 20 Minutes" in the cover letter while its MDHHS-6064 task table prints
"Total per month 62:21", and the task rows sum to 62:21 exactly.

**The task table wins:** it is the provider billing form, its own rows add up to it, and MDHHS has
paid against it. Taking the letter's figure billed a minute less than authorized on every invoice
for that client. Where the two disagree the import says so rather than choosing silently.

## day-grid-derivation

The MSA-1904 day grid is built from the **Number of Days** column, not by dividing Time-per-Month by
Time-per-Day. "2 days per week" in a 31-day month is 9 visits; dividing the times gave 8.

Tasks are anchored to weekdays so the sheet reads like a real schedule, one weekday pattern per task
group, walking forward a day per month. Travel time rides with the task it serves.

**Open issue:** the grid is *generated* from the authorization, and the owner's signature certifies
it. MDHHS describes the MSA-1904 as "an accurate record of the authorized Home Help services that
were provided on each day of the billing period". The sound fix is capturing actual caregiver visit
days. Raised 2026-09-01 and 2026-09-03; the owner is aware and has deferred it.

## email-subjects

No client identifier in any email subject, ever. A subject line sits unencrypted in server logs,
backups and phone notification previews.

- Invoices send the bare literal `INVOICE` from both the bulk and single paths.
- Document emails name the **form** — the app's own category, or "Document". Not the filename
  (files are routinely named after the client) and not a typed category (free text, and people type
  the client's name into it).

## roster-merge

What survives a background roster load: the server is authoritative for rows it returns; a **failed**
local save wins over it; a never-synced local addition is kept; a previously-synced row the server
no longer returns was deleted elsewhere and is dropped.

**Known consequence:** only unsynced rows survive an *empty* server response, so a transient empty
response empties the visible roster until the next load. The rows are safe on the server. Left
unchanged deliberately — altering sync semantics is how new bugs get made — but pinned by a test so
it is a decision rather than a surprise.

## backup-and-restore

Restore drops the concurrency token (`_rowVersion`) every record was backed up with. `expected_version`
means "the row I last read"; a restore has read nothing — it is a deliberate overwrite. Replaying the
backup's token made every record that still existed on the server reject the write, so a restore over
live rows wrote **nothing**, reported "NOT FULLY RESTORED — check your connection", and could never
succeed on retry. Restoring onto deleted rows always worked, which is why it hid.

Backups are also tagged by environment (`_staging`, `_local`). Every environment writes to the same
OneDrive folder and the automatic backup is named by date alone, so a staging session took
production's slot for the day — it happened on 2026-09-01.
