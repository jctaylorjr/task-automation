from collections import defaultdict

def find_printer_tasks(control_id_column, task_column, sheet_ranges):
    printer_tasks = defaultdict(list)

    for row in range(2, sheet_ranges.max_row + 1):
        control_id = sheet_ranges[f"{control_id_column}{row}"].value
        task = sheet_ranges[f"{task_column}{row}"].value
        if control_id is None or task is None:
            continue
        printer_tasks[control_id] = task

    return printer_tasks


def seen_printer(control_id_column, task_column, sheet_ranges):
    seen_printers = set()

    for row in range(2, sheet_ranges.max_row + 1):
        control_id = sheet_ranges[f"{control_id_column}{row}"].value
        task = sheet_ranges[f"{task_column}{row}"].value
        if control_id is not None and task is not None and control_id not in seen_printers:
            seen_printers.add(control_id)
        
    return seen_printers

def printer_workstation_mapping(control_id_column, workstation_column, sheet_ranges):
    printer_to_workstations = defaultdict(list)

    for row in range(2, sheet_ranges.max_row + 1):
        control_id = sheet_ranges[f"{control_id_column}{row}"].value
        workstation = sheet_ranges[f"{workstation_column}{row}"].value
        if control_id is None or workstation is None:
            continue
        printer_to_workstations[control_id].append(str(workstation))

    return printer_to_workstations
