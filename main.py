import time

from instruments.data_scrappers import DataScrapper
from instruments.import_generators import ImportGenerator
from instruments.data_instruments import DataInstruments
from instruments.resources import Resources


def main():
    DI.init_project() # Initialises all dirs and files for work

    # Chose options
    # region data_scrappers
    # DS.import_to_excel("name.xlsx", "name", 1)    # Useful for make table with names etc.(constructor)
    # DS.key_generator("name.xlsx")                 # Useful only for generating keys
    # DS.get_photo_data(True)                       # Create json dict file with photo links
    # endregion

    # region import_generators
    # IG.autofill_generator("filename.xlsx") # Universal
    # endregion

    # region data_instruments
    # DI.clean_descriptions()           # Cleans all data from all descriptions files
    # DI.description_splitter()         # Splits desctriptions from one file to descriptions dir
    # DI.how_many_marks()               # Simple print out number and all the marks
    # DI.check_duplicates()             # Simple print out all the duplicated marks (name + seat_year)
    # endregion


if __name__ == '__main__':
    start = time.time()

    DS, IG, DI, res = DataScrapper(), ImportGenerator(), DataInstruments(), Resources()
    main()
    res.close()

    print(f"\nTime elapsed: {time.time() - start} seconds")