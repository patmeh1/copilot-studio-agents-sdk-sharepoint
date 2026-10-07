"""Apply the SITE_ACCESS matrix (SharePoint Owners/Members/Visitors groups) using MarioR's (site owner) session."""
from cdp import session
from settings import PERSONAS, SHAREPOINT_ROOT, SITE_ACCESS, SITES
from signin import load_users
from sp import rest

GROUP = {"Owner": "AssociatedOwnerGroup", "Member": "AssociatedMemberGroup", "Visitor": "AssociatedVisitorGroup"}

if __name__ == "__main__":
    users = load_users()
    with session(new=True, port=PERSONAS["mario"]) as (b, ctx, page):
        for letter, path in SITES.items():
            site = SHAREPOINT_ROOT + path
            page.goto(site + "/_layouts/15/viewlsts.aspx", wait_until="domcontentloaded")
            page.wait_for_timeout(2500)
            for who, perms in SITE_ACCESS.items():
                level = perms.get(letter)
                if not level:
                    continue
                login = f"i:0#.f|membership|{users[who]['upn'].lower()}"
                status, _ = rest(page, "POST", f"/_api/web/{GROUP[level]}/users", {"LoginName": login}, site=site)
                print(f"Site {letter}: {who} -> {level}: HTTP {status}")
        page.close()
