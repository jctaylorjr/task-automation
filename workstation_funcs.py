from collections import defaultdict
from playwright.sync_api import Page

def get_workstation_lwsids(page: Page, ws_names):
    print("Looking up workstation LWSIDs...")
    ws_name_lwsids = []
    for ws in ws_names:
        if type(ws) is int:
            ws = f"W0{str(ws)}"
        # to make sure W0253604 and DIMFLH2664 can get in
        if len(ws) != 8 and len(ws) != 10:
            continue
        page.locator("#inputTxtWksta").fill(ws)
        with page.expect_response("https://partnershealthcare.service-now.com/api/now/sp/rectangle/ddd076c2473b39d02a94f147536d4327?id=esp_workstations") as response_info:
            page.locator("#wkstaSearch").click()
            response = response_info.value
            if response.ok:
                try:
                    ws_name_lwsid = response.json()["result"]["data"]["lws"]["searchList"][0]["displayValue"]
                except:
                    ws_name_lwsid = f"{ws} [Not found]"
                finally:
                    ws_name_lwsids.append(ws_name_lwsid)
    return ws_name_lwsids
