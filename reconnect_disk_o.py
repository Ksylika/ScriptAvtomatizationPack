import os
import time

drive_path = 'O:\\path\\toCheck\\netdisk'


def main(drive_path):
    while True:
        if not os.path.isdir(drive_path):
            os.system('net use /del o:')
            os.system('net use o: \\\\path\\to\\netdisk')
        else:
            pass
        time.sleep(300)


if __name__ == '__main__':
    main(drive_path)
