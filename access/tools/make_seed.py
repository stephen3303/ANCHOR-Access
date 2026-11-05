"""Write the data file's seed rows (backend/seed/*.tsv) from the ANCHOR web
app's mock data. Every message, person and unit in it is fictional.

The web app (read only, never edited from here) keeps its mock data in
src/lib/data.ts. Dump it to JSON once with Node 22 or later:

    node --experimental-strip-types dump.mts > data.json

(dump.mts imports data.ts and demo-users.ts and prints every exported value;
see the PR that added this script), then run

    python tools/make_seed.py data.json backend/seed

Each .tsv is UTF-8, one header row of column names (the table's field names),
then one row per record. In a value, a tab is written \\t, a line break \\n and
a backslash \\\\. An empty value is Null. Dates are yyyy-mm-dd hh:nn:ss (UTC, as
the web data has them).
"""
import json, sys, os, re

MONTHS = {m: i + 1 for i, m in enumerate("JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC".split())}


def esc(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "1" if v else "0"
    s = str(v)
    return s.replace("\\", "\\\\").replace("\t", "\\t").replace("\r\n", "\n").replace("\n", "\\n")


def iso(s):
    """2026-10-24T18:30:00Z -> 2026-10-24 18:30:00; 24 OCT 2026 -> 2026-10-24 00:00:00."""
    if not s:
        return None
    m = re.match(r"(\d{4}-\d\d-\d\d)T(\d\d:\d\d)(:\d\d)?", s)
    if m:
        return f"{m.group(1)} {m.group(2)}{m.group(3) or ':00'}"
    m = re.match(r"(\d{1,2}) ([A-Z]{3}) (\d{4})", s)
    if m:
        return f"{m.group(3)}-{MONTHS[m.group(2)]:02d}-{int(m.group(1)):02d} 00:00:00"
    raise ValueError("date? " + s)


def write(folder, name, cols, rows):
    with open(os.path.join(folder, name + ".tsv"), "w", encoding="utf-8", newline="\n") as f:
        f.write("\t".join(cols) + "\n")
        for r in rows:
            f.write("\t".join(esc(r.get(c)) for c in cols) + "\n")


def body_text(sections):
    """The official text: numbered paragraphs under each heading, as the web
    message view numbers them (1. PURPOSE. ...; a., b. for more than one)."""
    out = []
    for i, s in enumerate(sections, 1):
        paras = s["body"]
        if len(paras) == 1:
            out.append(f"{i}. {s['heading'].upper()}. {paras[0]}")
        else:
            out.append(f"{i}. {s['heading'].upper()}.")
            for j, p in enumerate(paras):
                out.append(f"    {chr(97 + j)}. {p}")
    return "\n\n".join(out)


def main(src, folder):
    d = json.load(open(src, encoding="utf-8"))
    os.makedirs(folder, exist_ok=True)
    msgs = d["MESSAGES"]
    key_of = {m["id"]: m["reference"] for m in msgs}

    write(folder, "tblMessages",
          ["MsgKey", "Dtg", "ReceivedOn", "Subject", "Reference", "MsgType", "Originator", "Source", "ToLine", "InfoLine",
           "Precedence", "Classification", "Importance", "Lifecycle", "RelatesTo", "ReplacedBy", "Communities",
           "AckRequired", "AckDue", "ActionText", "ActionSuspense", "ActionAssignedTo", "C2oixId", "Sha256", "SizeKb", "MsgClass",
           "Retention", "Distribution", "Related", "AiSummary", "RawText", "BodyText"],
          [{
              "MsgKey": m["id"], "Dtg": m["dtg"], "ReceivedOn": iso(m["received"]), "Subject": m["subject"],
              "Reference": m["reference"], "MsgType": m["type"], "Originator": m["originator"], "Source": m["source"],
              "ToLine": m["to"], "InfoLine": m["info"], "Precedence": m["precedence"],
              "Classification": m["classification"], "Importance": m["importance"], "Lifecycle": m["lifecycle"],
              "RelatesTo": m.get("relatesTo"), "ReplacedBy": m.get("replacedBy"),
              "Communities": "; ".join(m["communities"]), "AckRequired": m["ackRequired"],
              "AckDue": iso(m.get("ackDue")),
              "ActionText": (m.get("actionItem") or {}).get("text"),
              "ActionSuspense": (m.get("actionItem") or {}).get("suspense"),
              "ActionAssignedTo": (m.get("actionItem") or {}).get("assignedTo"),
              "C2oixId": m["c2oixId"], "Sha256": m["sha256"], "SizeKb": m["sizeKb"], "MsgClass": "A", "Retention": m["retention"],
              "Distribution": "; ".join(m["distribution"]),
              "Related": "; ".join(key_of.get(r, r) for r in m["related"]),
              "AiSummary": m.get("aiSummary"), "RawText": m["rawText"], "BodyText": body_text(m["sections"]),
          } for m in msgs])

    write(folder, "tblMessageSections", ["MsgKey", "SortOrder", "Heading", "Body"],
          [{"MsgKey": m["id"], "SortOrder": i, "Heading": s["heading"], "Body": "\n".join(s["body"])}
           for m in msgs for i, s in enumerate(m["sections"], 1)])
    write(folder, "tblAttachments", ["MsgKey", "FileName", "SizeText", "Kind"],
          [{"MsgKey": m["id"], "FileName": a["name"], "SizeText": a["size"], "Kind": a["kind"]}
           for m in msgs for a in m["attachments"]])
    write(folder, "tblAddressees", ["MsgKey", "AddrName", "AddrRole", "AddrKind"],
          [{"MsgKey": m["id"], "AddrName": a["name"], "AddrRole": a["role"], "AddrKind": a["kind"]}
           for m in msgs for a in m["addressees"]])
    write(folder, "tblAckReceipts", ["MsgKey", "CommandName", "ReceivedOn"],
          [{"MsgKey": k, "CommandName": r["command"], "ReceivedOn": iso(r["at"])}
           for k, rs in d["ACK_RECEIPTS"].items() for r in rs])
    write(folder, "tblGroupMembers", ["GroupName", "CommandName"],
          [{"GroupName": g["name"], "CommandName": c} for g in d["COMMAND_GROUPS"] for c in g["members"]])

    subs = d["SUBSCRIPTIONS"]
    seen = set()
    cat = []
    for s in subs + d["APPROVAL_QUEUE"]:
        if s["name"] in seen:
            continue
        seen.add(s["name"])
        c = s["criteria"]
        cat.append({"SubKey": s["id"], "SubName": s["name"], "Category": s["category"], "Description": s["description"],
                    "MsgTypes": "; ".join(c.get("messageTypes", [])), "Originators": "; ".join(c.get("originators", [])),
                    "Keywords": "; ".join(c.get("keywords", [])), "Precedences": "; ".join(c.get("precedence", [])),
                    "IsMandatory": s["mandatory"], "DefaultFrequency": s["frequency"], "IsActive": True})
    write(folder, "tblSubscriptions", ["SubKey", "SubName", "Category", "Description", "MsgTypes", "Originators",
                                       "Keywords", "Precedences", "IsMandatory", "DefaultFrequency", "IsActive"], cat)
    # The approval queue: other people's requests, waiting on an approver.
    write(folder, "tblPersonSubs", ["SubName", "PersonName", "Status", "Frequency", "ChInApp", "ChEmail", "ChTeams",
                                    "Justification", "RequestedOn", "LastUpdated"],
          [{"SubName": s["name"], "PersonName": s["requester"], "Status": "Pending Approval", "Frequency": s["frequency"],
            "ChInApp": s["channels"]["inApp"], "ChEmail": s["channels"]["email"], "ChTeams": s["channels"]["teams"],
            "Justification": s.get("justification"), "RequestedOn": iso(s["requestedOn"]),
            "LastUpdated": iso(s["requestedOn"])} for s in d["APPROVAL_QUEUE"]])

    write(folder, "tblAudit", ["LoggedAt", "UserName", "Action", "Resource", "Classification", "Result", "Device",
                               "Location", "Details"],
          [{"LoggedAt": iso(e["timestamp"]), "UserName": e["user"], "Action": e["action"], "Resource": e["resource"],
            "Classification": e["classification"], "Result": e["result"], "Device": e["device"],
            "Location": e["location"], "Details": e["details"]}
           for e in sorted(d["AUDIT_EVENTS"], key=lambda e: e["timestamp"])])

    write(folder, "tblWatchNotes", ["NoteText", "PersonName", "CreatedOn", "IsOpen"],
          [{"NoteText": n, "PersonName": "LT Taylor", "CreatedOn": "2026-10-24 18:00:00", "IsOpen": True}
           for n in d["WATCH_NOTES"]])
    write(folder, "tblOngoingTopics", ["Topic", "LastUpdate", "Status"],
          [{"Topic": t["topic"], "LastUpdate": iso(t["lastUpdate"]), "Status": t["status"]} for t in d["ONGOING_TOPICS"]])
    write(folder, "tblTurnoverChecklist", ["ItemText", "SortOrder", "IsActive"],
          [{"ItemText": t, "SortOrder": i, "IsActive": True} for i, t in enumerate(d["TURNOVER_CHECKLIST"], 1)])
    codes = {"Message Viewer": "VIEW", "Briefing Creator": "BRIEF", "Subscription Requester": "REQUEST",
             "Report Viewer": "REPORT", "Subscription Approver": "APPROVER", "System Admin": "ADMIN"}
    # Designated Release Authority belongs to drafting and release, which stay in Message Generator.
    write(folder, "tblRoles", ["RoleCode", "RoleName", "Detail", "SortOrder"],
          [{"RoleCode": codes[r["name"]], "RoleName": r["name"], "Detail": r["detail"], "SortOrder": i}
           for i, r in enumerate((r for r in d["CURRENT_USER"]["roles"] if r["name"] in codes), 1)])

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
