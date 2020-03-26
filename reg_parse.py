import yarp
import xls_writer
import argparse
from yarp import Registry

reg_file = open('NTUSER.DAT', 'rb')
reg = Registry.RegistryHive(reg_file)

docs = reg.find_key('SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs')


def main():
    # arg parser
    parser = argparse.ArgumentParser()
    # make argument flags
    parser.add_argument('-i', '--src', help='Source folder to scan', required=True, type=str)
    parser.add_argument('-o', '--dest', help='Output folder', required=True, type=str)
    # parse the arguments
    args = parser.parse_args()





    print(docs.last_written_timestamp())

    for i, value in enumerate(docs.subkeys()):
        print(f"{i}) {value.name()}, {value.values_count()}")


if __name__ == "__main__":
    main()