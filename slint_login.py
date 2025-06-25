import slint
import re
from tomlkit import parse, comment, document, nl, table, dumps
from os.path import isfile
from urllib.parse import unquote, unquote_plus

toml_file = "config.toml"

class Login(slint.loader.ui.login_window.MainWindow):

    def __init__(self):
        super().__init__()
        self.toml = None
        if isfile(toml_file) == True:
            with open(toml_file, "r") as file:
                self.toml = parse(file.read())
            self.load_toml()
        self.username = ""
        self.password = ""

    @slint.callback
    def load_toml(self):
        self.username_input = self.toml["username"]

    @slint.callback
    def attempt_login(self, username, password):
        
        self.username = username
        self.password = password
        self.write_toml_file()
        self.hide()
    
    def write_toml_file(self):
        with open(toml_file, "w") as f:
            if self.toml is None:
                self.toml = document()
                self.toml.add("username", self.username) 
                f.write(dumps(self.toml))
            else:
                self.toml["username"] = self.username
                f.write(dumps(self.toml))
            
