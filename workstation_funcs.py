from collections import defaultdict
from playwright.sync_api import Page

def print_workstations(workstations):
    if len(workstations) > 1:
        print(" ".join(workstations))

def printer_workstation_mapping(control_id_column, workstation_column, sheet_ranges):
    printer_to_workstations = defaultdict(list)

    for row in range(2, sheet_ranges.max_row + 1):
        control_id = sheet_ranges[f"{control_id_column}{row}"].value
        workstation = sheet_ranges[f"{workstation_column}{row}"].value
        if control_id is None or workstation is None:
            continue
        printer_to_workstations[control_id].append(str(workstation))

    return printer_to_workstations

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

def seen_printer(control_id_column, task_column, sheet_ranges):
    seen_printers = set()

    for row in range(2, sheet_ranges.max_row + 1):
        control_id = sheet_ranges[f"{control_id_column}{row}"].value
        task = sheet_ranges[f"{task_column}{row}"].value
        if control_id is not None and task is not None and control_id not in seen_printers:
            seen_printers.add(control_id)
        
    return seen_printers