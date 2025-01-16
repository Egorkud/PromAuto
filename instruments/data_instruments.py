from pathlib2 import Path
from openpyxl.workbook import Workbook

from instruments import config

def init_project():
    def create_path(*files):
        for i in files:
            if not i.exists():
                if i.suffix:
                    i.touch(exist_ok=True)
                    print(f"File '{i}' created")
                else:
                    i.mkdir(exist_ok=True)
                    print(f"Directory '{i}' created")

    # data dir
    folder_path = Path("data")
    excel_path = folder_path / "Empty file for auto seat case.xlsx"
    json_path = folder_path / "links_data.json"

    create_path(folder_path, json_path)

    if not excel_path.exists():
        wb = Workbook()
        products_sheet = wb.active

        products_sheet.title = "Export Products Sheet"
        for id, name in config.PRODUCTS_COLUMNS.items():
            products_sheet.cell(1, id).value = name

        groups_sheet = wb.create_sheet("Export Groups Sheet")
        for id, name in config.GROUPS_COLUMNS.items():
            groups_sheet.cell(1, id).value = name

        wb.save(excel_path)
        print(f"File '{excel_path}' created")

    # Descriptions dir
    folder_path = Path("descriptions")
    description_file_ru = folder_path / "Description main ru.txt"
    description_file_ukr = folder_path / "Description main ukr.txt"
    folder_ru = folder_path / "ru"
    folder_ukr = folder_path / "ukr"
    models_ru = folder_ru / "models"
    models_ukr = folder_ukr / "models"

    create_path(folder_path, description_file_ru, description_file_ukr,
                folder_ru, folder_ukr, models_ru, models_ukr)

    for i in range(1, 31):
        create_path(folder_ru / f"{i}.txt", folder_ukr / f"{i}.txt")
    for i in range(1, 11):
        create_path(models_ru / f"{i}.txt", models_ukr / f"{i}.txt")

    # Lists of cars, Work result dir
    lists_of_cars = Path("lists of cars")
    work_result = Path("work result")
    create_path(lists_of_cars, work_result)