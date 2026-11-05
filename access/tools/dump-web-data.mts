import * as d from "./data.ts";
import * as u from "./demo-users.ts";
const out:any = {};
for (const [k,v] of Object.entries(d)) if (typeof v !== "function") out[k]=v;
out.actionItems = d.actionItems();
out.DEMO_USERS = u.DEMO_USERS;
console.log(JSON.stringify(out,null,1));
