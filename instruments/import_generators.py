from instruments import data_instruments as DI

def autofill_generator(filename, models_sheet, blank_book, blank_products, blank_groups, groups_sheet, data_changes, links_data):
    colours = data_changes["colours"]
    start_position = 0
    dif_number = len(data_changes["colours"])
    description_count = data_changes["Кількість описів"]
    dcounter = 1
    counter = start_position
    groups_data = DI.get_groups_data(groups_sheet)
    print(groups_data)
    DI.fulfill_groups_data(blank_groups, groups_data, data_changes["Parent group"])

    for row in range(counter + 2, models_sheet.max_row * dif_number + 2, dif_number):
        # region Variables initialisation
        name_engl = models_sheet.cell(start_position + 1, column=2).value.strip()
        name_ru = models_sheet.cell(start_position + 1, column=3).value.strip()
        name_ukr = models_sheet.cell(start_position + 1, column=4).value.strip()
        seat_year_ru = models_sheet.cell(start_position + 1, column=5).value
        seat_year_ukr = models_sheet.cell(start_position + 1, column=6).value

        mark = models_sheet.cell(start_position + 1, column=7).value
        model = models_sheet.cell(start_position + 1, column=8).value
        series = models_sheet.cell(start_position + 1, column=9).value
        year = models_sheet.cell(start_position + 1, column=10).value
        compatibility = models_sheet.cell(start_position + 1, column=11).value

        price = models_sheet.cell(start_position + 1, column=12).value
        group_id = models_sheet.cell(start_position + 1, column=13).value
        group_name = models_sheet.cell(start_position + 1, column=14).value

        key_ru = models_sheet.cell(start_position + 1, column=15).value
        key_ukr = models_sheet.cell(start_position + 1, column=16).value

        note = models_sheet.cell(start_position + 1, column=17).value

        colours_sequence_ru = data_changes["colours"]
        colours_sequence_ukr = data_changes["colours_ukr"]

        titles = ("Название_Характеристики", "Измерение_Характеристики", "Значение_Характеристики")
        # endregion

        for i in range(dif_number):
            colour = colours[i]
            # 1. Артикул (A)
            blank_products.cell(counter + row + i, column=1).value = f"{data_changes["Артикул"][0]}{data_changes["Артикул"][1] + start_position}"

            # region 2. Назва ру-англ, укр-англ (B, C)
            if data_changes["first_name_colour"]:
                colours_seq_dif_ru = [i for i in colours]
                colours_seq_dif_ukr = [i for i in colours_sequence_ukr]
            else:
                colours_seq_dif_ru = [""] + [i for i in colours[1:]] # + [""] for no colour main position
                colours_seq_dif_ukr = [""] + [i for i in colours_sequence_ukr[1:]]

            blank_products.cell(counter + row + i, column=2).value = DI.replace_names(data_changes["name_ru"], name_ru, name_engl, seat_year_ru, colours_seq_dif_ru[i], data_changes["names_with_colours"])
            blank_products.cell(counter + row + i, column=3).value = DI.replace_names(data_changes["name_ukr"], name_ukr, name_engl, seat_year_ukr, colours_seq_dif_ukr[i], data_changes["names_with_colours"])
            # endregion

            # region 3. Поисковые запросы ru (D)
            blank_products.cell(counter + row + i, column=4).value = key_ru

            # Поисковые запросы ukr (E)
            blank_products.cell(counter + row + i, column=5).value = key_ukr
            # endregion

            # region 4. Описи
            # Описи однакові (F)
            if data_changes["Кількість описів"] == 1:
                languages = ("ru", "ukr")
                names = (name_ru, name_ukr)
                for index, language in enumerate(languages):
                    with open(f"Descriptions/Description main {language}.txt", "r", encoding="utf-8") as file:
                        data = file.read()
                        new_data = data.replace("name", f"{names[index]}")  # Із додаванням назви
                        blank_products.cell(counter + row, column=6 + index).value = new_data

            # Описи різні (F)
            else:
                new_name = name_engl # blank_sheet.cell(counter + row, column=3).value put this inside, used it before
                blank_products.cell(counter + row, column=6).value = DI.descriptions_generator(dcounter, new_name, "ru")
                blank_products.cell(counter + row, column=7).value = DI.descriptions_generator(dcounter, new_name, "ukr")
            # endregion

            # region 5. Дефолтні характеристики (ніколи не змінював)
            # Валюта (J)
            blank_products.cell(counter + row + i, column=10).value = data_changes["Валюта"]

            # Наличие (P)
            blank_products.cell(counter + row + i, column=16).value = data_changes["Наличие"]

            # Адрес подраздела (T)
            blank_products.cell(counter + row + i, column=20).value = data_changes["Адрес подраздела"]

            # Идентификатор подраздела (AA)
            blank_products.cell(counter + row + i, column=27).value = data_changes["Идентификатор подраздела"]

            # ID группы разновидностей (AF)
            blank_products.cell(counter + row + i, column=32).value = start_position + 1

            # Цена от (AK)
            blank_products.cell(counter + row + i, column=37).value = data_changes["Цена от"]
            # endregion

            # region 6. Основні характеристики (до 50 колонки)
            # Ціна (I), якщо не задати одну ціну, будуть попередні з файлу даних
            blank_products.cell(counter + row + i, column=9).value = price if data_changes["Цена"] is True else data_changes["Цена"]

            # Единица (K)
            blank_products.cell(counter + row + i, column=11).value = data_changes["Комплектация"]

            # Изображение, ссылки (O)
            try:
                blank_products.cell(counter + row + i, column=15).value = links_data[mark][colour]
            except KeyError:
                blank_products.cell(counter + row + i, column=15).value = links_data[colour]
            except TypeError:
                if note == "main2":
                    full_text = f"{links_data[f"{mark}2"]}, {links_data['additional_photos']}"
                else:
                    full_text = f"{links_data[mark]}, {links_data['additional_photos']}"
                blank_products.cell(counter + row + i, column=15).value = full_text

            # Номер группы (по необходимости) (R)
            if data_changes["New groups"]:
                blank_products.cell(counter + row + i, column=28).value = groups_data[mark]
            else:
                blank_products.cell(counter + row + i, column=18).value = groups_data[mark]

            # Производитель (AC)
            blank_products.cell(counter + row + i, column=29).value = data_changes["Производитель"]

            # Страна производства (AD)
            blank_products.cell(counter + row + i, column=30).value = data_changes["Страна_производства"]

            # Личные заметки вверху (AG), щоб тільки головному різновиду давати, але можливо без різниці
            blank_products.cell(counter + row + i, column=33).value = data_changes["Личные_заметки"]
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

            for key, value in data_changes["Дополнительные_характеристики"].items():
                if type(value) is bool:
                    if value:
                        value = new_data[key]
                    else:
                        continue

                blank_products.cell(counter + row + i, count_chars).value = key
                blank_products.cell(counter + row + i, count_chars + 2).value = value
                for index, title in enumerate(titles):
                    blank_products.cell(1, count_chars + index).value = title
                count_chars += 3
            blank_products.cell(1, count_chars).value = "ID_Сопутствующих"
            # endregion

        dcounter += 1
        if dcounter > description_count:
            dcounter = 1

        print(f"{start_position + 1}/{models_sheet.max_row}")
        start_position += 1

    print(f"File created: {filename}")
    blank_book.save(filename)
