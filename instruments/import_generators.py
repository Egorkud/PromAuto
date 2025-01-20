from instruments.data_instruments import DataInstruments


class ImportGenerator(DataInstruments):
    def __init__(self):
        super().__init__()

    def autofill_generator(self, filename):
        colours = self.data_changes["colours"]
        start_position = 0
        dif_number = len(self.data_changes["colours"])
        description_count = self.data_changes["Кількість описів"]
        dcounter = 1
        counter = start_position
        groups_data = self.get_groups_data()
        print(groups_data)
        self.fulfill_groups_data(groups_data, self.data_changes["Parent group"])

        for row in range(counter + 2, self.models_sheet.max_row * dif_number + 2, dif_number):
            # region Variables initialisation
            name_engl = self.models_sheet.cell(start_position + 1, column=2).value.strip()
            name_ru = self.models_sheet.cell(start_position + 1, column=3).value.strip()
            name_ukr = self.models_sheet.cell(start_position + 1, column=4).value.strip()
            seat_year_ru = self.models_sheet.cell(start_position + 1, column=5).value
            seat_year_ukr = self.models_sheet.cell(start_position + 1, column=6).value

            mark = self.models_sheet.cell(start_position + 1, column=7).value
            model = self.models_sheet.cell(start_position + 1, column=8).value
            series = self.models_sheet.cell(start_position + 1, column=9).value
            year = self.models_sheet.cell(start_position + 1, column=10).value
            compatibility = self.models_sheet.cell(start_position + 1, column=11).value

            price = self.models_sheet.cell(start_position + 1, column=12).value
            group_id = self.models_sheet.cell(start_position + 1, column=13).value
            group_name = self.models_sheet.cell(start_position + 1, column=14).value

            key_ru = self.models_sheet.cell(start_position + 1, column=15).value
            key_ukr = self.models_sheet.cell(start_position + 1, column=16).value

            note = self.models_sheet.cell(start_position + 1, column=17).value

            colours_sequence_ru = self.data_changes["colours"]
            colours_sequence_ukr = self.data_changes["colours_ukr"]

            titles = ("Название_Характеристики", "Измерение_Характеристики", "Значение_Характеристики")
            # endregion

            for i in range(dif_number):
                colour = colours[i]
                # 1. Артикул (A)
                self.blank_products.cell(counter + row + i, column=1).value = f"{self.data_changes["Артикул"][0]}{self.data_changes["Артикул"][1] + start_position}"

                # region 2. Назва ру-англ, укр-англ (B, C)
                if self.data_changes["first_name_colour"]:
                    colours_seq_dif_ru = [i for i in colours]
                    colours_seq_dif_ukr = [i for i in colours_sequence_ukr]
                else:
                    colours_seq_dif_ru = [""] + [i for i in colours[1:]] # + [""] for no colour main position
                    colours_seq_dif_ukr = [""] + [i for i in colours_sequence_ukr[1:]]

                self.blank_products.cell(counter + row + i, column=2).value = self.replace_names(self.data_changes["name_ru"], name_ru, name_engl, seat_year_ru, colours_seq_dif_ru[i], self.data_changes["names_with_colours"])
                self.blank_products.cell(counter + row + i, column=3).value = self.replace_names(self.data_changes["name_ukr"], name_ukr, name_engl, seat_year_ukr, colours_seq_dif_ukr[i], self.data_changes["names_with_colours"])
                # endregion

                # region 3. Поисковые запросы ru (D)
                self.blank_products.cell(counter + row + i, column=4).value = key_ru

                # Поисковые запросы ukr (E)
                self.blank_products.cell(counter + row + i, column=5).value = key_ukr
                # endregion

                # region 4. Описи
                # Описи однакові (F)
                if self.data_changes["Кількість описів"] == 1:
                    languages = ("ru", "ukr")
                    names = (name_ru, name_ukr)
                    for index, language in enumerate(languages):
                        with open(f"Descriptions/Description main {language}.txt", "r", encoding="utf-8") as file:
                            data = file.read()
                            new_data = data.replace("name", f"{names[index]}")  # Із додаванням назви
                            self.blank_products.cell(counter + row, column=6 + index).value = new_data

                # Описи різні (F)
                else:
                    new_name = name_engl # blank_sheet.cell(counter + row, column=3).value put this inside, used it before
                    self.blank_products.cell(counter + row, column=6).value = self.descriptions_generator(dcounter, new_name, "ru")
                    self.blank_products.cell(counter + row, column=7).value = self.descriptions_generator(dcounter, new_name, "ukr")
                # endregion

                # region 5. Дефолтні характеристики (ніколи не змінював)
                # Валюта (J)
                self.blank_products.cell(counter + row + i, column=10).value = self.data_changes["Валюта"]

                # Наличие (P)
                self.blank_products.cell(counter + row + i, column=16).value = self.data_changes["Наличие"]

                # Адрес подраздела (T)
                self.blank_products.cell(counter + row + i, column=20).value = self.data_changes["Адрес подраздела"]

                # Идентификатор подраздела (AA)
                self.blank_products.cell(counter + row + i, column=27).value = self.data_changes["Идентификатор подраздела"]

                # ID группы разновидностей (AF)
                self.blank_products.cell(counter + row + i, column=32).value = start_position + 1

                # Цена от (AK)
                self.blank_products.cell(counter + row + i, column=37).value = self.data_changes["Цена от"]
                # endregion

                # region 6. Основні характеристики (до 50 колонки)
                # Ціна (I), якщо не задати одну ціну, будуть попередні з файлу даних
                self.blank_products.cell(counter + row + i, column=9).value = price if self.data_changes["Цена"] is True else self.data_changes["Цена"]

                # Единица (K)
                self.blank_products.cell(counter + row + i, column=11).value = self.data_changes["Комплектация"]

                # Изображение, ссылки (O)
                try:
                    self.blank_products.cell(counter + row + i, column=15).value = self.links_data[mark][colour]
                except KeyError:
                    self.blank_products.cell(counter + row + i, column=15).value = self.links_data[colour]
                except TypeError:
                    if note == "main2":
                        full_text = f"{self.links_data[f"{mark}2"]}, {self.links_data['additional_photos']}"
                    else:
                        full_text = f"{self.links_data[mark]}, {self.links_data['additional_photos']}"
                    self.blank_products.cell(counter + row + i, column=15).value = full_text

                # Номер группы (по необходимости) (R)
                if self.data_changes["New groups"]:
                    self.blank_products.cell(counter + row + i, column=28).value = groups_data[mark]
                else:
                    self.blank_products.cell(counter + row + i, column=18).value = groups_data[mark]

                # Производитель (AC)
                self.blank_products.cell(counter + row + i, column=29).value = self.data_changes["Производитель"]

                # Страна производства (AD)
                self.blank_products.cell(counter + row + i, column=30).value = self.data_changes["Страна_производства"]

                # Личные заметки вверху (AG), щоб тільки головному різновиду давати, але можливо без різниці
                self.blank_products.cell(counter + row + i, column=33).value = self.data_changes["Личные_заметки"]
                # endregion

                # region 7. Характеристики та моделі, обрані сайтом
                count_chars = 50
                new_data = {
                    "Цвет": colour,
                    "Марка": mark,
                    "Модель": model,
                    "Серия": series,
                    "Год выпуска автомобиля": year,
                } # Ініціалізація даних

                for key, value in self.data_changes["Дополнительные_характеристики"].items():
                    if type(value) is bool:
                        if value:
                            value = new_data[key]
                        else:
                            continue

                    self.blank_products.cell(counter + row + i, count_chars).value = key
                    self.blank_products.cell(counter + row + i, count_chars + 2).value = value
                    for index, title in enumerate(titles):
                        self.blank_products.cell(1, count_chars + index).value = title
                    count_chars += 3
                self.blank_products.cell(1, count_chars).value = "ID_Сопутствующих"
                # endregion

            dcounter += 1
            if dcounter > description_count:
                dcounter = 1

            print(f"{start_position + 1}/{self.models_sheet.max_row}")
            start_position += 1

        print(f"File created: {filename}")
        self.blank_book.save(filename)
