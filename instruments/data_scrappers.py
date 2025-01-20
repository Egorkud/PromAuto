import json

from instruments import config
from instruments.data_instruments import DataInstruments


class DataScrapper(DataInstruments):
    def __init__(self):
        super().__init__()

    def import_to_excel(self, name, parent_group_name, step):
        # Create parent group if necessary
        duplicates_groups = [None, parent_group_name]
        self.new_groups_sheet.cell(1, 1).value = 1
        self.new_groups_sheet.cell(1, 2).value = parent_group_name

        for row in range(1, self.export_sheet.max_row, step):
            self.empty_sheet.cell(row, column=1).value = row  # Новий артикул (просто номер)

            data_ru = self.export_sheet.cell(row + 1, column=2).value.split(";")
            data_ukr = self.export_sheet.cell(row + 1, column=3).value.split(";")
            name_engl = data_ru[0].strip()
            name_ru = data_ru[3].strip()
            name_ukr = data_ukr[3].strip()

            # region Задання років, або типів авто (седан...)
            try:
                seats_ru = f"{data_ru[1].strip()} {data_ru[2].strip()}"
                seats_ukr = f"{data_ukr[1].strip()} {data_ukr[2].strip()}"
            except Exception as ex:
                seats_ru = None
                seats_ukr = None
                print(ex)
            # endregion

            # region Імена, вторинні характеристики, тут зазвичай нічого не змінюємо
            self.empty_sheet.cell(row, column=2).value = name_engl
            self.empty_sheet.cell(row, column=3).value = name_ru
            self.empty_sheet.cell(row, column=4).value = name_ukr

            self.empty_sheet.cell(row, column=5).value = seats_ru
            self.empty_sheet.cell(row, column=6).value = seats_ukr

            mark = model = series = year = compatibility = None
            for i in range(50, 85):
                cell = self.export_sheet.cell(row + 1, column=i).value
                if cell == "Марка":
                    mark = self.export_sheet.cell(row + 1, column=i + 2).value
                elif cell == "Модель" or cell == "Мoдель":
                    model = self.export_sheet.cell(row + 1, column=i + 2).value
                elif cell == "Серия":
                    series = self.export_sheet.cell(row + 1, column=i + 2).value
                elif cell == "Год выпуска автомобиля":
                    year = self.export_sheet.cell(row + 1, column=i + 2).value
                elif cell == "Совместимость":
                    compatibility = self.export_sheet.cell(row + 1, column=i + 2).value

            self.empty_sheet.cell(row, column=7).value = mark
            self.empty_sheet.cell(row, column=8).value = model
            self.empty_sheet.cell(row, column=9).value = series
            self.empty_sheet.cell(row, column=10).value = year
            self.empty_sheet.cell(row, column=11).value = compatibility
            self.empty_sheet.cell(row, column=12).value = self.export_sheet.cell(row + 1, column=9).value # Price
            # endregion

            # region Групи
            if mark is None:
                self.empty_sheet.cell(row, column=13).value = 1
                print("\nMARK IS NONE")
            else:
                group_id, group_name, duplicates_groups = self.create_group(mark, duplicates_groups)
                self.new_groups_sheet.cell(group_id, 1).value = group_id
                self.new_groups_sheet.cell(group_id, 2).value = group_name
                self.new_groups_sheet.cell(group_id, 3).value = group_name
                self.new_groups_sheet.cell(group_id, 4).value = 1    # Parend group id

                self.empty_sheet.cell(row, column=13).value = group_id
                self.empty_sheet.cell(row, column=14).value = group_name
            # endregion

            # region Ключові запити
            # Отримання ключів із експорту
            key_ru = self.export_sheet.cell(row + 1, column=4).value
            key_ukr = self.export_sheet.cell(row + 1, column=5).value

            # Подарок водителю добавить
            key_ru, keys_ukr = self.add_gift_keys(key_ru, key_ukr, name_ru, name_ukr)

            self.empty_sheet.cell(row, column=15).value = key_ru
            self.empty_sheet.cell(row, column=16).value = key_ukr
            #endregion

            # region Додаткова інфо (нотатки) 17 колонка
            # Personal conditions add here to column 17
            # endregion

            print(row)

        print(f"File created: {name}")
        self.book_empty.save(name)

    def key_generator(self, name):
        full_keys_name_ru = config.keys_ru
        full_keys_name_ukr = config.keys_ukr
        for row in range(1, self.models_sheet.max_row + 1):
            self.empty_sheet.cell(row, column=1).value = row

            name_engl = self.models_sheet.cell(row, column=2).value
            name_ru = self.models_sheet.cell(row, column=3).value
            name_ukr = self.models_sheet.cell(row, column=4).value

            # engl_orig = original
            # engl_big1 = first letter is big, other are small
            # ru_big1 = first letter is big, other are small
            # ru_small = all letters are small

            new_keys_ru = full_keys_name_ru.replace("engl_orig", f"{name_engl}")
            new_keys_ru = new_keys_ru.replace("engl_big1", f"{name_engl.lower().title()}")
            new_keys_ru = new_keys_ru.replace("ru_orig", f"{name_ru}")
            new_keys_ru = new_keys_ru.replace("ru_small", f"{name_ru.lower()}")

            new_keys_ukr = full_keys_name_ukr.replace("engl_orig", f"{name_engl}")
            new_keys_ukr = new_keys_ukr.replace("engl_big1", f"{name_engl.lower().title()}")
            new_keys_ukr = new_keys_ukr.replace("ukr_orig", f"{name_ukr}")
            new_keys_ukr = new_keys_ukr.replace("ukr_small", f"{name_ukr.lower()}")

            self.empty_sheet.cell(row, column=2).value = new_keys_ru
            self.empty_sheet.cell(row, column=3).value = new_keys_ukr

        print(f"File created: {name}")
        self.book_empty.save(name)

    def get_photo_data(self, colours_for_one_mark):
        colours = self.data_changes["colours"]
        marks = self.get_all_marks()
        if colours_for_one_mark:
            models_dict = self.create_empty_coloured_dict(marks, colours)
        else:
            models_dict = self.create_empty_marks_coloured_dict(marks, colours)

        for row in range(2, self.export_sheet.max_row + 1):
            link = self.export_sheet.cell(row, 15).value
            mark = self.get_mark(row)
            colour = self.get_colour(row)

            if colours_for_one_mark and len(colours) == 1:
                models_dict[mark] = link
            elif colours_for_one_mark:
                models_dict[colour] = link
            else:
                models_dict[mark][colour] = link
            # print(models_dict)
        with open("data/links_data.json", "w", encoding="utf-8") as file:
            file.write(json.dumps(models_dict, indent=4))
