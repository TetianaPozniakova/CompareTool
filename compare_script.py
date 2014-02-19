import argparse
import time

from Comparator import Comparator


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("reportFolder_path", help="Path to the folder with reports for comparison.")
    parser.add_argument("-c", "--compareScript_folder", nargs='?', default="C:\\CompareScripts\\",
                        help="Path to the folder with compare scripts.")
    args = parser.parse_args()
    start_time = time.time()
    Comparator(args.reportFolder_path, args.compareScript_folder).compare()
    print (time.time() - start_time)/60, "minutes"