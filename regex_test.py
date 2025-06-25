import re

s = "Zebra ZD611D-HC Desktop Direct Thermal Printer - Monochrome - Label Print - Fast Ethernet - USB - US"

match = re.search(r"^[^-]*", s, re.IGNORECASE)
if match:
    print(match.group(0))