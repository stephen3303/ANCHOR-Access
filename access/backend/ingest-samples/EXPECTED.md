# Mail ingest test set

Ten synthetic messages for the Admin page's **Mail ingest** tab (front end
`modAnchorIngest`). Every command, serial and message here is fictional and
UNCLASSIFIED; the classification markings are demonstration only. None of the
serials collides with the seed data (`COMSUBLANT 078/26`, `NAVADMIN 236/26` and
the like): the senders are the fictional NAVSTA KESTREL and SUBGRU TEAL.

## How to run them

1. Make a fresh data file (`backend\new-backend.ps1` with the seed folder) and
   build the front end against it; sign in as a System Admin.
2. Copy all ten `.txt` files (not this file) into the drop folder: by default
   the `ingest` folder beside the data file (Admin > Mail ingest shows the path
   under DROP FOLDER; **Check folder now** makes it when it's missing).
3. Admin > Mail ingest > **Check folder now**. The files are read in name
   order, so 02 is stored before 05 cancels it and 01 before 03 and 04 repeat it.
4. Expect the toast and **Last run**: `4 stored, 1 repeat, 1 conflict, 4 rejected`.
   The drop folder is empty; `Ingested\` holds 01, 02, 03, 05 and 09;
   `Rejected\` holds 04, 06, 07, 08 and 10, each with a `.reason.txt`.
5. Run **Check folder now** again: `nothing waiting`. Copy 01 back in and check
   again: `1 repeat`, and `Ingested\` gets `01-clean-kestrel-101 (2).txt`.

## Expected results

| File | Result | Why / what to check |
|---|---|---|
| 01-clean-kestrel-101.txt | stored | NAVSTA KESTREL 101/26, ROUTINE, type FLEET MSG (a GENADMIN whose serial names no known type), class A, UNCLASSIFIED. Needs acknowledgement, due 07 NOV 1200Z. Sections Purpose, Action, Enclosures (two items) and Point of Contact. Addressees: All COMSUBLANT Units as a group, USS EXAMPLEFISH (SSN-990) and SUBGRU TEAL as commands. Related: NAVSTA KESTREL PORT OPS GUIDE. LF line ends. |
| 02-clean-subgru-teal-014-optask.txt | stored | SUBGRU TEAL 014/26, PRIORITY, type OPTASK (named in MSGID). Saved as UTF-8 with a byte-order mark. ACTIVE until 05 runs. |
| 03-repeat-kestrel-101.txt | repeat | The same text as 01 with CRLF line ends: same serial, same SHA-256 once the line breaks are made the same. Logged; nothing stored; the file goes to Ingested. Double-click its row: NAVSTA KESTREL 101/26 opens. |
| 04-conflict-kestrel-101.txt | conflict | Serial NAVSTA KESTREL 101/26 again, but paragraph 2 says 111200Z instead of 101200Z: same serial, different hash. Refused and logged; the file goes to Rejected. |
| 05-cancels-subgru-teal-014.txt | stored | SUBGRU TEAL 015/26, IMMEDIATE, with NARR "THIS MSG CANCELS REF A" and REF A = SUBGRU TEAL 014/26. After the run SUBGRU TEAL 014/26 is CANCELLED, replaced by SUBGRU TEAL 015/26 (its message page shows "Cancelled by SUBGRU TEAL 015/26"). The log's reason says so. |
| 06-bad-dtg.txt | rejected | `"321530Z NOV 26" is not a valid date-time group.` (day 32). |
| 07-unknown-type.txt | rejected | `unknown type: WIDGETREP is not an active message type here...` Add WIDGETREP under Admin > Message types, copy the file from Rejected back to the drop folder, check again: stored as NAVSTA KESTREL 104/26, type WIDGETREP. |
| 08-missing-subj.txt | rejected | `Missing the SUBJ/ set.` |
| 09-secret-demo-marking.txt | stored | SUBGRU TEAL 016/26 marked SECRET (the spaced "S E C R E T" line; the "DEMONSTRATION MARKING ONLY" line under it is passed over). Lists and the message page show the demonstration marking beside SECRET. With HIGHEST MARKING set to CONFIDENTIAL or lower it is rejected instead. |
| 10-above-secret.txt | rejected | `The text is marked above SECRET. ANCHOR refuses it and stores nothing.` Neither the serial nor the marking goes into the log or the audit trail. |

Every attempt, whatever its result, adds one row to the ingest log and one
"Ingest" entry to the audit trail (Audit page; Success for stored, Warning for
repeat, Failed for conflict and rejected). Verify on the Audit page still
passes afterwards.

## The mailbox feeder

The review VMs have no Outlook profile, so **Check mailbox now** there should
only say that Outlook isn't available (or that the mailbox isn't set) and
change nothing. With Outlook: set MAILBOX and MAIL FOLDER, Save, mail the
files to that mailbox as `.txt` attachments (or paste one as the mail's text),
then **Check mailbox now**. Each mail moves to an Ingested or Rejected folder
under the mail folder, and the log row carries its Internet Message-ID.
