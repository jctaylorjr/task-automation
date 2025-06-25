import slint
import re
from tomlkit import parse, document, dumps
from os.path import isfile

toml_file = "config.toml"

class Controller(slint.loader.ui.app_window.MainWindow):

    def __init__(self):
        super().__init__()
        self.toml = None
        if isfile(toml_file) == True:
            with open(toml_file, "r") as file:
                self.toml = parse(file.read())
            self.load_toml()

    @slint.callback
    def load_toml(self):
        self.username = self.toml["username"]
        self.excel_link = self.toml["excel_link"]
        self.sheet_name = self.toml["sheet_name"]
        self.start_row = self.toml["start_row"]
        self.end_row = self.toml["end_row"]
        self.location_default = self.toml["location_default"]
        self.room_column = self.toml["room_column"]
        self.entity_column = self.toml["entity_column"]
        self.control_id_column = self.toml["control_id_column"]
        self.task_column = self.toml["task_column"]
        self.epic_dep_column = self.toml["epic_dep_column"]
        self.workstation_column = self.toml["workstation_column"]

    @slint.callback
    def attempt_login(self):
        self.write_toml_file()

        self.sharepoint_group = re.search(r"(.*?)\/Shared", self.excel_link).group(1)
        
        self.excel_file_path = re.search(r"(\/sites\/.*?xlsx)", self.excel_link).group(1)

        # print(f"unqoute: {unquote(self.excel_file_path)}")
        # print(f"unquote plus: {unquote_plus(self.excel_file_path)}")
        # excel_file = os.path.basename(unquote(self.excel_file_path))
        # print(f"File name: {excel_file}")

        # print("matches\n\n")
        # self.excel_file_name = re.findall(r"\/(.*?xlsx)", el)
        # for match in self.excel_file_name:
        #     print(match)
        # print("\n\n")

        self.hide()
    
    def write_toml_file(self):
        with open(toml_file, "w") as f:
            if self.toml is None:
                self.toml = document()
                self.toml.add("username", self.username) 
                self.toml.add("excel_link", self.excel_link)
                self.toml.add("sheet_name", self.sheet_name)
                self.toml.add("start_row", self.start_row)
                self.toml.add("end_row", self.end_row)
                self.toml.add("location_default", self.location_default)
                self.toml.add("room_column", self.room_column)
                self.toml.add("entity_column", self.entity_column)
                self.toml.add("control_id_column", self.control_id_column)
                self.toml.add("task_column", self.task_column)
                self.toml.add("epic_dep_column", self.epic_dep_column)
                self.toml.add("workstation_column", self.workstation_column)
                f.write(dumps(self.toml))
            else:
                self.toml["username"] = self.username
                self.toml["excel_link"] = self.excel_link
                self.toml["sheet_name"] = self.sheet_name
                self.toml["start_row"] = self.start_row
                self.toml["end_row"] = self.end_row
                self.toml["location_default"] = self.location_default
                self.toml["room_column"] = self.room_column
                self.toml["entity_column"] = self.entity_column
                self.toml["control_id_column"] = self.control_id_column
                self.toml["task_column"] = self.task_column
                self.toml["epic_dep_column"] = self.epic_dep_column
                self.toml["workstation_column"] = self.workstation_column
                f.write(dumps(self.toml))
