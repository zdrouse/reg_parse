# Python Project 2
# Zack Rouse
# CSI 5740
# A python script to parse recent document data from registry hives.

# native
import logging
import argparse

# 3rd party
import xlsxwriter
from yarp import Registry

reg_file = open('NTUSER.DAT', 'rb')
reg = Registry.RegistryHive(reg_file)
recent_docs = reg.find_key('SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs')

def export_filenames(recent_docs):
    for i, value in enumerate(recent_docs.subkeys()):
        #print(f"{i}) {value.name()}, {value.values_count()}")
        print(value.name())
        k = recent_docs.subkey(value.name())
        for j, thing in enumerate(k.values()):
            try:
                if thing.name() != "MRUListEx":
                # filename will provide a parsed string of the filename from REG_BINARY raw data output from yarp. Without this, we have ugly raw data.
                    filename = thing.data()[::2][:thing.data()[::2].find(b'\x00')].decode()
                    print(f"    {filename}")
            except:
                print(f"No data for: {thing.name()}")

def export_filetypes(recent_docs):
    for i, value in enumerate(recent_docs.subkeys()):
        print(f"{i}) {value.name()}, {value.values_count()}")

def main():
    # arg parser
    parser = argparse.ArgumentParser()
    # make argument flags
    parser.add_argument('-i', '--src', help='Source folder to scan', required=True, type=str)
    parser.add_argument('-o', '--dest', help='Output folder', required=True, type=str)
    # parse the arguments
    args = parser.parse_args()

    export_filenames(recent_docs)


if __name__ == "__main__":
    main()