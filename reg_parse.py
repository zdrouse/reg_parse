# Python Project 2
# Zack Rouse
# CSI 5740
# A python script to parse recent document data from registry hives.

# native
import logging
import argparse
import os

# 3rd party
import xlsxwriter
from yarp import Registry


def get_registry_hive(input_file):
    reg_file = open(input_file, 'rb')
    reg = Registry.RegistryHive(reg_file)
    return reg

def export_filenames(recent_docs, out_path):
    try:
        with open(f'{out_path}\\results.log', 'w+') as f:
            for i, value in enumerate(recent_docs.subkeys()):
                k = recent_docs.subkey(value.name())
                print(value.name())
                f.write(value.name() + "\n")
                for j, thing in enumerate(k.values()):
                    try:
                        # Remove the empty items
                        if thing.name() != "MRUListEx":
                            # filename will provide a parsed string of the REG_BINARY raw data output from yarp. Without this, we have ugly raw binary data.
                            filename = thing.data()[::2][:thing.data()[::2].find(b'\x00')].decode()
                            print(f"    {filename}")
                            f.write(f"    {filename}" + "\n")
                    except:
                        print(f"No data for: {thing.name()}")
            f.close()
    except Exception as e:
        print(f"Error exporting filename log: {e}")
        exit(1)

def export_filetypes(recent_docs, out_path):
    try:
        workbook = xlsxwriter.Workbook(f'{out_path}\\results.xlsx')
        worksheet = workbook.add_worksheet()
        row = 1
        col = 0
        worksheet.write(0, 0, 'TYPE')
        worksheet.write(0, 1, 'COUNT')
        file_extensions = ['.docx', '.xlsx', '.pptx', '.avi', '.jpg']
        for i, value in enumerate(recent_docs.subkeys()):
            if value.name() in file_extensions:
                worksheet.write(row, col, value.name())
                worksheet.write(row, col + 1, value.values_count())
                row += 1
        workbook.close()
    except Exception as e:
        print(f"Error exporting filetypes excel file: {e}")
        exit(1)

def main():
    # arg parser
    parser = argparse.ArgumentParser()
    # make argument flags
    parser.add_argument('-i', '--src', help='Source folder to scan', required=True, type=str)
    parser.add_argument('-o', '--dest', help='Output folder', required=True, type=str)
    # parse the arguments
    args = parser.parse_args()

    try:
        registry_hive = get_registry_hive(args.src)
        recent_docs = registry_hive.find_key('SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs')
    except Exception as e:
        print(f"Could not get registry hive, please verify the input provided: {e}")
        exit(1)

    export_filenames(recent_docs, args.dest)
    export_filetypes(recent_docs, args.dest)

if __name__ == "__main__":
    main()