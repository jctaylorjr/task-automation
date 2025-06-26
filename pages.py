from playwright.sync_api import Page
from typing import Optional
import re

class ExcelPage:
    def __init__(self, page: Page):
        self.page = page

    def load(self, excel_url, excel_sheet, username):
        url = excel_url
        self.page.goto(url)
        self.page.get_by_role("textbox", name="someone@example.com").click()
        self.page.get_by_role("textbox", name="someone@example.com").fill(username)
        self.page.get_by_role("button", name="Next").click()
        try:
            self.page.wait_for_url(url)
        except:
            print("Didn't wait for excel page url...", end="")
        try:
            self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("tab", name=f"{excel_sheet}").click()
        except:
            print("failed to click excel_tab")

    def set_task(self, column, row, task):
        task_url = f"https://partnershealthcare.service-now.com/now/nav/ui/classic/params/target/sc_task.do%3Fsys_id%3D{task}"
        self.go_to(column, row)
        self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.locator("#gridKeyboardContentEditable_textElement").press("ControlOrMeta+k")
        try:
            self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("textbox", name="Display Text").click(timeout=5000)
        except:
            self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.locator("#gridKeyboardContentEditable_textElement").press("ControlOrMeta+k")
            self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("textbox", name="Display Text").click()
                
        self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("textbox", name="Display Text").fill(task)
        self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("textbox", name="Type or paste URL").click()
        self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("textbox", name="Type or paste URL").fill(task_url)
        self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("button", name="OK").click()

    def go_to(self, column: str, row: str):
        self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("combobox", name="Name Box").click(force=True)
        self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("combobox", name="Name Box").fill(f"{column}{row}", force=True)
        self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("combobox", name="Name Box").press("Enter")
    
    # def get_cell_value(self, column: str, row: int) -> Optional[str]:
    #     try:
    #         self.go_to(column, row)
    #         return self.page.locator("iframe[name=\"WacFrame_Excel_0\"]").content_frame.get_by_role("textbox", name="formula bar").input_value(timeout=5000)
    #     except:
    #         print("failed to get cell value")
    #         return None

    # def check_for_task(self, row: int) -> Optional[str]:
    #     task = self.get_cell_value(task_column, row)
    #     if task == "" or task is None:
    #         task = None
    #     return task
        

class SearchPage:
    def __init__(self, page: Page):
        self.page = page

    def load(self):
        self.page.goto("https://partnershealthcare.service-now.com/esp")
        self.page.wait_for_url(
            "https://partnershealthcare.service-now.com/esp", timeout=300000
        )
        self.page.get_by_role("link", name="Printer Inventory Have your").click()

    def search(self, control_id, asset_tag):
        self.page.get_by_role("searchbox", name="Type here to search").click()
        self.page.get_by_role("searchbox", name="Type here to search").fill(
            f"{control_id}"
        )
        try:
            #xb17afece47ae21102a94f147536d4346 > div > div.div-background > div.panel.panel-default > div:nth-child(4) > a > h5
            #xb17afece47ae21102a94f147536d4346 > div > div.div-background > div.panel.panel-default > div:nth-child(4) > a
            self.page.locator(".div-background > div.div-background > div.panel.panel-default > div.list-group.ng-scope > a > h5").get_by_text(f"{control_id} ({asset_tag})").click()
            # strict mode violation
            # self.page.locator(".div-background > div.div-background > div.panel.panel-default > div.list-group.ng-scope > a > h5").get_by_text(printer.control_id).wait_for(timeout=1000, state='visible')
        except:
            try:
                self.page.locator(".div-background > div.div-background > div.panel.panel-default > div.list-group.ng-scope > a > h5").get_by_text(f"{control_id} ({asset_tag})").click(delay=10000)
            except:
                print("Couldn't find printer search result anchors")
                print("Make sure all fields are filled out before unpausing")
                self.page.pause()
    
    def fill_required_fields(self, control_id, location, room, entity, epic_dep, printer_model):
        self.page.wait_for_timeout(2000)
        try:
            # LOCATION FILLED HERE
            print("Filling location...", end=" ", flush=True)
            self.page.locator("#s2id_location > a > span.select2-arrow").click(force=True)
            self.page.get_by_role("combobox", name="*Location", exact=True, disabled=False).fill(f"{location}")
            self.page.get_by_role("option", name=f"{location}").first.click()
        except:
            self.page.pause()
        else:
            try:
                print(f"{self.page.locator("#select2-chosen-4").inner_text(timeout=1000)}")
            except:
                print(f"{location}")


        
        # ROOM FILLED HERE
        print("Filling Room/Cube...", end=" ", flush=True)
        self.page.get_by_role("textbox", name="*Room/Cube").click(force=True)
        self.page.get_by_role("textbox", name="*Room/Cube", disabled=False).fill(f"{room}")
        print(f"{room}")
        

        # ENTITY FILLED HERE
        # self.page.get_by_role("link", name="BWH | Brigham and Women's Clear field x_pahcs_edm_epic_entity").click()
        # self.page.get_by_role("combobox", name="*Epic Entity", exact=True).click()
        # if self.page.locator("#s2id_epicEntity > a > abbr").is_visible():
        #     self.page.locator("#s2id_epicEntity > a > abbr").click()
        try:
            print("Filling Epic Entity...", end=" ", flush=True)
            self.page.locator("#s2id_epicEntity > a > span.select2-arrow").click(force=True)
            self.page.get_by_role("combobox", name="*Epic Entity", exact=True, disabled=False).fill(f"{entity}")
            self.page.get_by_role("option", name=f"{entity}").first.click(timeout=3000)
        except:
            self.page.pause()
        else:
            try:
                print(f"{self.page.locator("#select2-chosen-6").inner_text(timeout=1000)}")
            except:
                print(f"{entity}")

        # self.page.get_by_role("combobox", name="*Epic Entity", exact=True).press("Enter", delay=2000)
        # self.page.get_by_role("link", name="BWH PRE/PACU (10030010422) Clear field x_pahcs_edm_epic_dep").click()


        # EPIC DEP FILLED HERE
        # self.page.get_by_role("combobox", name="Epic DEP", exact=True).click()
        # if self.page.locator("#s2id_epicDepartment > a > abbr").is_visible():
        #     self.page.locator("#s2id_epicDepartment > a > abbr").click()
        try:
            print("Filling Epic DEP...", end=" ", flush=True)
            self.page.locator("#s2id_epicDepartment > a > span.select2-arrow").click(force=True)
            self.page.get_by_role("combobox", name="Epic DEP", exact=True, disabled=False).fill(f"{epic_dep}")
            self.page.get_by_role("option", name=f"{epic_dep}").first.click(timeout=3000)
        except:
            self.page.pause()
        else:
            try:
                print(f"{self.page.locator("#select2-chosen-8").inner_text(timeout=1000)}")
            except:
                print(f"{epic_dep}")
        # self.page.get_by_role("combobox", name="Epic DEP", exact=True).press("Enter", delay=2000)


        # PRINTER MODEL FILLED HERE
        # self.page.get_by_role("link", name="Zebra ZD611 Clear field").click()
        # self.page.get_by_role("combobox", name="*Epic Printer Model", exact=True).click()
        # if self.page.locator("#s2id_epicModel > a > abbr").is_visible():
        #     self.page.locator("#s2id_epicModel > a > abbr").click()
        try:
            print("Filling Epic Printer Model...", end=" ", flush=True)
            self.page.locator("#s2id_epicModel > a > span.select2-arrow").click(force=True)
            self.page.get_by_role("combobox", name="*Epic Printer Model", exact=True, disabled=False).fill(f"{printer_model}")
            self.page.get_by_role("option", name=f"{printer_model}").first.click(timeout=3000)
        except:
            self.page.pause()
        else:
            try:
                print(f"{self.page.locator("#select2-chosen-9").inner_text(timeout=1000)}")
            except:
                print(f"{printer_model}")

        # self.page.get_by_role("combobox", name="*Epic Printer Model", exact=True).press("Enter", delay=2000)

        # updates printer button, wait for success banner to confirm update went through and closes banner to not trigger again when searching for another printer
        try:
            #xfb6f0bc6476661102a94f147536d431a > div > div > div.div-container-no-border > p > input
            # self.page.locator("div > div > div.div-container-no-border > p > input").click()
            print("Updating printer configuration...")
            self.page.get_by_role("button", name="Update Printer Configuration").click(force=True, delay=500)
            locator = self.page.locator("#uiNotificationContainer > div > span").get_by_text("updated!")
            locator.wait_for(state="visible")
            for button in self.page.get_by_role("button", name="Close Notification").all():
                button.click()
        except:
            # print("couldnt wait for config ID invalid reference!")
            print("Please check order page to ensure that the printer configuration was updated. Press green arrow in debug menu to continue task creation...")
            self.page.pause()

    def get_epic_entity(self) -> Optional[str]:
        entity = None
        try:
            entity = self.page.locator("#s2id_currentEpicEntity").get_by_role("link").input_value()
            # self.page.locator('stack div-container-no-border').filter(has_text='Current Epic Entity').wait_for('span', timeout=1000)
        except:
            print("Couldn't find Current Epic Entity span")
        # entity = (
        # self.page.locator('stack div-container-no-border').filter(has_text='Current Epic Entity').locator('span').input_value()
        # # self.page.get_by_label("Current Epic Entity").
        # # .content_frame.get_by_role("textbox", name="Serial number")
        # # .input_value())
        # )
        finally:
            print(f"Epic entity: {entity}")
            return entity

class OrderPage:
    def __init__(self, page: Page):
        self.page = page

    def load(self):
        self.page.goto(
            "https://partnershealthcare.service-now.com/now/nav/ui/classic/params/target/home.do"
        )

    def order_task(self, control_id, entity, serial_number, workstations) -> str:
        self.reset_to_new_request()

        # Pressing order ticket button and copying TASK#
        print("Ordering ticket...", end=" ", flush=True)
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "button", name="Order Now"
        ).click(delay=1000)
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "link", name=re.compile(r"TASK[0-9]+", re.IGNORECASE)
        ).click()
        task = (
            self.page.locator('iframe[name="gsft_main"]')
            .content_frame.get_by_role("textbox", name="Number")
            .input_value()
        )
        print(task)
        
        # Entering control ID into ticket request form
        print(f"Entering Configuration item...", end=" ", flush=True)
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "button", name="EPR and LRS Request"
        ).click()
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "searchbox", name="Mandatory - must be populated"
        ).click()
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "combobox", name="Mandatory - must be populated"
        ).fill(f"{control_id}")
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "combobox", name="Mandatory - must be populated"
        ).press("Enter")
        print(control_id)
        
        # Entering filled out description with entity, control id, and serial number
        print(f"Filling description...")
        map_request = ""
        if len(workstations) > 0:
            map_request = f"\nPlease map the following workstations to this printer: {", ".join(workstations)}"
        description = f"Request Epic EPR and LRS build for the following Beaker Specimen Label Printer:\nEntity: {entity}\nControl #: {control_id}\nSerial #: {serial_number}\n*To be configured for Wi-Fi and Hostname will need to be updated via PPME{map_request}"
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "textbox",
            name=" Field value has changed since last updateDescription  Search Knowledge To",
        ).click()
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "textbox",
            name=" Field value has changed since last updateDescription  Search Knowledge To",
        ).fill(description)
        print(description)
        print("description filled")

        # Waits for config ID invalid reference to go away before updating ticket
        try:
            locator = self.page.locator("iframe[name=\"gsft_main\"]").content_frame.get_by_label("Catalog Task form section").get_by_text("Invalid reference")
            locator.wait_for(state="hidden")
        except:
            print("couldnt wait for config ID invalid reference!")
            print("Please check order page to ensure that the task was updated and submitted. Press green arrow in debug menu to continue task creation...")
            self.page.pause()
        else:
            self.page.locator('iframe[name="gsft_main"]').content_frame.locator(
                "#sysverb_update"
            ).click()
            print(f"{task} updated!")
        finally:
            return task


    def reset_to_new_request(self):
        self.page.get_by_role("menuitem", name="All").click()
        self.page.get_by_role("link", name="Service Catalog 1 of").click()
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "combobox", name="Search catalog"
        ).click()
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "combobox", name="Search catalog"
        ).fill("other (request)")
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "combobox", name="Search catalog"
        ).press("Enter")
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "link", name="other (request)"
        ).click()
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "combobox", name="   Assignment Group"
        ).click()
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "combobox", name="   Assignment Group"
        ).fill("Beaker Technicians - MGB")
        self.page.locator('iframe[name="gsft_main"]').content_frame.get_by_role(
            "combobox", name="   Assignment Group"
        ).press("Enter")

class PrinterSerialSearch:
    def __init__(self, page: Page):
        self.page = page
        self.current_cid = 0

    def load(self):
        self.page.goto(
            "https://partnershealthcare.service-now.com/now/nav/ui/classic/params/target/home.do"
        )
        self.page.wait_for_url(
            "https://partnershealthcare.service-now.com/now/nav/ui/classic/params/target/home.do", timeout=0
        )

    def goto_printer_cmdb(self, control_id):
        try:
            url = f"https://partnershealthcare.service-now.com/now/nav/ui/classic/params/target/cmdb_ci_printer.do%3Fsys_id%3D{control_id}"
            self.page.goto(url)
        except:
            print("Failed to goto page printer page with control ID")
        else:
            try:
                self.page.wait_for_url(url)
            except:
                print("Didn't wait for printer CMDB url")

    def find_serial_number(self) -> Optional[str]:
        sn = None
        try:
            sn = (
            self.page.locator('iframe[name="gsft_main"]')
            .content_frame.get_by_role("textbox", name="Serial number")
            .input_value(timeout=3000))
            if sn == "":
                sn = None
        except:
            print("Failed to get serial number from printer ServiceNow page")
        finally:
            # print(f"Serial number from ServiceNow: {sn}")
            return sn

    def find_asset_tag(self) -> Optional[str]:
        asset_tag = None
        try:
            asset_tag = self.page.locator("iframe[name=\"gsft_main\"]").content_frame.get_by_role("textbox", name="Asset tag").input_value(timeout=3000)
            # asset_tag = (
            # self.page.locator('iframe[name="gsft_main"]')
            # .content_frame.get_by_role("textbox", name="Asset tag")
            # .input_value(timeout=5000))
            if asset_tag == "":
                asset_tag = None
        except:
            print("Failed to get Asset tag from printer ServiceNow page")
        finally:
            # print(f"Asset tag from ServiceNow: {asset_tag}")
            return asset_tag

    def find_printer_model(self) -> Optional[str]:
        # "#cmdb_ci_printer.model_id_label"
        printer_model = None
        try:
            printer_model = self.page.locator("iframe[name=\"gsft_main\"]").content_frame.get_by_role("textbox", name="Read only - cannot be modifiedModel ID").input_value(timeout=3000)
            # printer_model = (self.page.locator('#cmdb_ci_printer.model_id_label').input_value(timeout=5000))
            if printer_model == "":
                printer_model = None
        except:
            print("Failed to get Printer model from printer ServiceNow page")
        finally:
            # print(f"Printer model from ServiceNow: {printer_model}")
            return printer_model



def flush_print(s: str, end="", flush=True) -> None:
    print(f"{s}", end=end, flush=flush)