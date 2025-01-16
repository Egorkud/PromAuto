import time

from pathlib import Path
from openpyxl.workbook import Workbook

from instruments import config
from instruments import data_instruments as DI
from instruments import data_scrappers as DS
from instruments import  import_generators as IG


def main():
    start = time.time()
    DI.init_project() # Initialises all dirs and files for work



    print(f"\nTime elapsed: {time.time() - start} seconds")

if __name__ == '__main__':
    main()