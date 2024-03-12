import re
import pandas as pd
from pathlib import Path
import warnings
import argparse


TRACEBACK = False

warnings.filterwarnings("ignore")


class bcolors:
    OKBLUE = "\033[94m"
    OKCYAN = "\033[96m"
    OKGREEN = "\033[92m"
    WARNING = "\033[93m"
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    UNDERLINE = "\033[4m"


suffixes = [
    ".xlsx",
]


def find_excel(path):
    return [
        x
        for x in Path(path).rglob("*")
        if x.suffix in suffixes and not x.stem.startswith("$~")
    ]


def read_excel(excel_path):
    return pd.read_excel(excel_path, sheet_name=None)


def color_match(row, word):
    hits = list(re.finditer(word, row))

    new_string = []
    last_end = 0

    for hit in hits:
        start, stop = hit.span()
        new_string.append(row[last_end:start])
        new_string.append(bcolors.FAIL)
        new_string.append(bcolors.UNDERLINE)
        new_string.append(row[start:stop])
        new_string.append(bcolors.ENDC)
        last_end = stop

    new_string.append(row[last_end:])

    return "".join(new_string)


def find_word_in_excel(excel_file, word):
    excel_df = read_excel(excel_file)

    for sheet in excel_df.keys():
        df = excel_df[sheet].to_numpy()

        for i, row in enumerate(df):
            row_str = " ".join(str(x) for x in row)
            if re.findall(word, row_str, flags=re.IGNORECASE):
                print(
                    f"{bcolors.UNDERLINE}{bcolors.OKCYAN}File: {excel_file.resolve().parts[-2]}/{excel_file.resolve().parts[-1]} "
                    f"| {bcolors.UNDERLINE}{bcolors.OKBLUE}Sheet: {sheet} "
                    f"| {bcolors.UNDERLINE}{bcolors.OKGREEN}Line {i + 1}{bcolors.ENDC}"
                )
                row_str = color_match(row_str, word)
                print(row_str)


def run(folder, word):
    excel_files = find_excel(folder)

    for excel in excel_files:
        try:
            find_word_in_excel(excel, word)
        except Exception as e:
            print(
                f"   {bcolors.UNDERLINE}{bcolors.FAIL}[ERROR]: File did not work {excel}"
            )
            print(f"   {bcolors.WARNING}Reason: {e}")
            if TRACEBACK:
                import traceback

                print(traceback.format_exc())


def main():
    parser = argparse.ArgumentParser(
        description="Search for a word in Excel files within a folder."
    )
    parser.add_argument("folder", help="The folder path to search for Excel files.")
    parser.add_argument("word", help="The word to search for within Excel files.")
    args = parser.parse_args()

    print(
        f"{bcolors.OKGREEN}========== SEARCHING {args.folder} =========={bcolors.ENDC}"
    )

    run(args.folder, args.word)

    print(f"{bcolors.OKGREEN}========== DONE =========={bcolors.ENDC}")


if __name__ == "__main__":
    main()
