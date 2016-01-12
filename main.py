from main_workbook import main_workbook_all_check
from master_check import master_check
import sys, os

def main():
    main_workbook_all_check()
    print ('\nMaster-check Complete')
    os.system('pause')
    sys.exit()

if __name__ == "__main__":
    main()
