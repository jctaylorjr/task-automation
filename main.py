from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File 
from slint_app import Controller
from urllib.parse import unquote
from pages import ExcelPage, PrinterSerialSearch, OrderPage, SearchPage
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
import os
import re
from workstation_funcs import get_workstation_lwsids, format_workstation, is_valid_workstation
from mapping import printer_workstation_mapping, seen_printer, find_printer_tasks, map_printers_workstations_tasks
import math
from datetime import datetime, timezone, timedelta
from typing import Optional
import socket

sharepoint_url = "https://partnershealthcare.sharepoint.com/sites/IPEDProceduralSpecimenCollectionHardwareWorkgroup"
# https://partnershealthcare-my.sharepoint.com/:x:/r/personal/ktran36_mgb_org/Documents/Attachments/Net%20New%20Beaker%20Printers_EPRLRS_TASKS_NEEDED.xlsx?d=wa7e951e28f6d4ad0bada14345a7508f5&csf=1&web=1&e=OaLFHw

def main():
    c = Controller()
    c.run()

    # try:
    #     ctx_auth = AuthenticationContext(c.sharepoint_group, browser_mode=True)
    #     if ctx_auth.acquire_token_for_user(c.username, c.password):
    #         ctx = ClientContext(sharepoint_url, ctx_auth)
    #         web = ctx.web
    #         ctx.load(web)
    #         ctx.execute_query()

        # ctx_auth.with_credentials(user_credentials)
    #     # https://partnershealthcare.sharepoint.com/sites/IPEDProceduralSpecimenCollectionHardwareWorkgroup/Shared%20Documents/General/Data%20Collection%20&%20TDR/BWH/BWH_Label%20Printer%20Data%20Collection%20Sheet_Logical%20Mapping.xlsx?web=1

    # except:
    #     print("Failed to login, closing...")
    # else:
    #     print('Authenticated into sharepoint as: ',web.properties['Title'])

    if os.path.isfile("task_log.csv") != True:
        with open('task_log.csv', mode='a') as file:
            file.write("Date,Printer Control ID,Task,Serial Number,Asset Tag,Printer Model,Epic Entity,Location,Room,Department,Workstations,Workbook,Sheet,Row\n")


    # excel_file = os.path.basename(unquote(c.excel_file_path))
    # print(f"File name: {excel_file}")

    # response = File.open_binary(ctx, c.excel_file_path)
    excel_link = "https://partnershealthcare-my.sharepoint.com/:x:/r/personal/ktran36_mgb_org/Documents/Attachments/Net%20New%20Beaker%20Printers_EPRLRS_TASKS_NEEDED.xlsx?d=wa7e951e28f6d4ad0bada14345a7508f5&csf=1&web=1&e=OaLFHw"
    excel_file = "newbeakers.xlsx"
    # with open(excel_file, 'wb') as output_file:  
    #     output_file.write(response.content)

    wb = load_workbook(excel_file)
    wb.active
    sheet_ranges = wb[c.sheet_name]

    # seen, mappings, printer_tasks = map_printers_workstations_tasks(c.control_id_column, c.workstation_column, c.task_column, sheet_ranges)

    # mappings = printer_workstation_mapping(c.control_id_column, c.workstation_column, sheet_ranges)
    # seen = seen_printer(c.control_id_column, c.task_column, sheet_ranges)
    # printer_tasks = find_printer_tasks(c.control_id_column, c.task_column, sheet_ranges)

    timezone_offset = -4.0  # EST (UTC−04:00) Daylight Savings time 
    tzinfo = timezone(timedelta(hours=timezone_offset))

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        context.set_default_timeout(0)

        search_page = SearchPage(context.new_page())
        excel_page = ExcelPage(context.new_page())
        order_page = OrderPage(context.new_page())
        serial_number_search_page = PrinterSerialSearch(context.new_page())
        # ws_esp = context.new_page()

        search_page.load()
        # excel_page.load(c.excel_link, c.sheet_name, c.username)
        excel_page.load(excel_link, c.sheet_name, c.username)
        order_page.load()
        serial_number_search_page.load()
        # ws_esp.goto("https://partnershealthcare.service-now.com/esp?id=esp_workstations")

        for row in range(int(c.start_row), int(c.end_row) + 1):
            print(f"\n{row}")
            control_id = sheet_ranges[f"{c.control_id_column}{row}"].value
            if control_id is None or type(control_id) is not int or cid_len(control_id) != 6:
                print("Control ID None, not int, not length of 6 skipping...")
                continue
            task = sheet_ranges[f"{c.task_column}{row}"].value
            # if task is not None:
            #     continue

            # workstation = sheet_ranges[f"{c.workstation_column}{row}"].value
            # workstation = format_workstation(workstation)
            # if is_valid_workstation(workstation) is False:
            #     print("Not a valid workstation, skipping...\n\n")
            #     continue

            # if location is None:
            #     print("Not a valid location, skipping...\n\n")
            #     continue
            print("Searching control ID on servicenow...", end=" ", flush=True)
            serial_number_search_page.goto_printer_cmdb(control_id)
            serial_number = serial_number_search_page.find_serial_number()
            if serial_number is None:
                print("Incomplete info, skipping...")
                continue
            asset_tag = serial_number_search_page.find_asset_tag()
            printer_model = serial_number_search_page.find_printer_model()
            print(f"Serial number: {serial_number}, Asset tag: {asset_tag}, Printer model: {printer_model}...", end=" ", flush=True)
            if serial_number is None or asset_tag is None or printer_model is None:
                print("Incomplete info, skipping...")
                continue
            else:
                print("", flush=True)

            # try:
            #     printer_model = printer_model_regex(printer_model)
            #     if printer_model != "Zebra ZD611D":
            #         print(printer_model)
            #         # continue
            #     elif printer_model == "Zebra ZD611D":
            #         print(f"Changing printer model {printer_model} to Zebra ZD611")
            #         printer_model = "Zebra ZD611"
            # except:
            #     print("Couldnt parse printer model")
            #     # continue
                
            if serial_number is not None and asset_tag is not None:
                settings = search_page.search(control_id, asset_tag)
                if settings is None:
                    print("Failed to get json from printer search")
                    continue
                # fields = search_page.fill_required_fields(control_id, location, room, entity, epic_dep, printer_model)
            entity, room, epic_dep, location, printer_model = settings 

            update = False
            if entity is None or entity == "":
                print(f"Entity: {entity}, skipping...")
                continue
                entity = sheet_ranges[f"{c.entity_column}{row}"].value
            if room is None or room == "":
                print(f"Room: {room}, skipping...")
                continue
                # room = sheet_ranges[f"{c.room_column}{row}"].value
            if epic_dep is None or epic_dep == "":
                epic_dep = search_page.fill_epic_dep(sheet_ranges[f"{c.epic_dep_column}{row}"].value)
                if epic_dep is not None:
                    update = True
                else:
                    epic_dep = ""
                # epic_dep = sheet_ranges[f"{c.epic_dep_column}{row}"].value
            if location is None or location == "":
                print(f"Location: {location}, skipping...")
                continue
                location = sheet_ranges[f"{c.location_column}{row}"].value
            if printer_model is None or printer_model == "":
                print(f"Printer model: {printer_model}, skipping...")
                continue

            if update is True:
                search_page.update_config()

            # if task is None and control_id is not None and entity is not None and room is not None and epic_dep is not None and workstation is not None:
            print(f"row {row}: Entity: {entity}, Control ID: {control_id}, Room/cube: {room}, Department: {epic_dep}, Task: {task}")
            # if control_id in seen:
            #     print("Task already made for this printer...")
            #     task = printer_tasks[control_id]
            #     print("Filling in associated ticket...", end="", flush=True)
            #     excel_page.set_task(c.task_column, row, task)
            #     continue

            # try:
            #     printer_model = printer_model_regex(printer_model)
            #     if printer_model != "Zebra ZD611D":
            #         print(printer_model)
            #         continue
            #     elif printer_model == "Zebra ZD611D":
            #         print(f"Changing printer model {printer_model} to Zebra ZD611")
            #         printer_model = "Zebra ZD611"
            # except:
            #     print("Couldnt parse printer model")
            #     continue

                # workstations = get_workstation_lwsids(ws_esp, mappings[control_id])
                # print(workstations)
            workstations = []
            if task is None:
                task = order_page.order_task(control_id, entity, serial_number, workstations)
                with open('task_log.csv', mode='a') as file:
                    file.write(f"{datetime.now(tzinfo)},{control_id},{task},{serial_number},{asset_tag},{printer_model},{entity},{location},{room},{epic_dep},{workstations},{excel_file},{c.sheet_name},{row}\n")
                # sheet_ranges[f"{c.task_column}{row}"].hyperlink = f"https://partnershealthcare.service-now.com/now/nav/ui/classic/params/target/sc_task.do%3Fsys_id%3D{task}"
                # sheet_ranges[f"{c.task_column}{row}"].value = f"{task}"
                excel_page.set_task(c.task_column, row, task)
            excel_page.set_cell(c.location_column, row, location)
            # sheet_ranges[f"{c.location_column}{row}"] = location
            excel_page.set_cell(c.epic_dep_column, row, epic_dep)
            # sheet_ranges[f"{c.epic_dep_column}{row}"] = epic_dep
            excel_page.set_cell(c.room_column, row, room)
            # sheet_ranges[f"{c.room_column}{row}"] = room
            excel_page.set_cell("H", row, serial_number)
            # sheet_ranges[f"{"H"}{row}"] = serial_number
            # hostnames = [f"csc{serial_number}", serial_number, f"csc{control_id}"]
            # for hostname in hostnames:
            #     result = resolve_hostname(hostname)
            #     if not all(result):
            #         continue
            #     else:
            #         hn, ip = result
            #         excel_page.set_cell("I", row, hn)
            #         excel_page.set_cell("J", row, ip)
            # seen.add(control_id)

        # os.remove(excel_file)
        context.close()
        # try:
        #     wb.save()
        # except:
        #     wb.save(filename=excel_file)
    
    # # df = pd.read_excel("https://partnershealthcare.sharepoint.com/sites/IPEDProceduralSpecimenCollectionHardwareWorkgroup/Shared%20Documents/General/Data%20Collection%20&%20TDR/BWH/BWH_Label%20Printer%20Data%20Collection%20Sheet_Logical%20Mapping.xlsx", "BWH IP_ED_Proc") 

# class Printer:
#     def __init__(self, row, cid_col, task_col, entity_col):
#         self.control_id = row[f"{cid_col}{row}"].value
#         self.task = row[f"{task_col}{row}"].value
#         self.entity = row[f"{entity_col}{row}"].value
#         self.serial_number

def resolve_hostname(hostname: str) -> Optional[tuple]:
    try:
        ipv4 = socket.gethostbyname(hostname)
        return hostname, ipv4
    except Exception as e:
        print(e)
    finally:
        return None

def printer_model_regex(s: str):
    match = re.search(r"^[^-]*", s, re.IGNORECASE)
    if match:
        # print(f"regex match: {match.group(0)}")
        return match.group(0)
    else:
        return None

def cid_len(n):
    try:
        return int(math.log10(n))+1
    except:
        return 0

if __name__ == "__main__":
    main()
