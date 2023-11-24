import pandas as pd
from bs4 import BeautifulSoup, Comment
import re
from tabulate import tabulate
import argparse
import os
import sys

def excel_to_dict(file_path, col_search, col_replace):
    df = pd.read_excel(file_path, sheet_name=0, header=None)
    data_dict = {str(k).strip(): str(v).strip() for k, v in zip(df[int(col_search)], df[int(col_replace)]) if pd.notna(k) and pd.notna(v)}
    return data_dict

def replace_text_in_html(input_file, output_file, replace_dict):
    replaced_count = {old_text: 0 for old_text in replace_dict.keys()}

    with open(input_file, 'r', encoding='utf-8') as file:
        html_content = file.read()

    soup = BeautifulSoup(html_content, 'html.parser')

    for old_text, new_text in replace_dict.items():
        pattern = re.compile(re.escape(old_text), re.IGNORECASE)
        for tag in soup.find_all(string=True):
            if tag.parent.name not in ['style', 'script', 'head', 'title', 'meta', '[document]']:
                if not isinstance(tag, Comment):
                    matches = re.findall(pattern, tag)
                    if matches:
                        count = len(matches)
                        replaced_count[old_text] += count
                        tag.replace_with(re.sub(pattern, new_text, tag))

    out_html = soup.encode(formatter="html", encoding='utf-8')  # Specify UTF-8 encoding
    with open(output_file, 'w', encoding='utf-8') as file:  # Specify UTF-8 encoding
        file.write(out_html.decode('utf-8'))

    return replaced_count


def sort_dict_by_key_length(data_dict):
    sorted_dict = dict(sorted(data_dict.items(), key=lambda item: len(item[0]), reverse=True))
    return sorted_dict


def print_replacements_summary(replaced_count):
    colorized_data = []

    for old_text, count in replaced_count.items():
        color_code = '\033[92m' if count > 0 else '\033[91m'  # ANSI escape code for colors
        reset_color = '\033[0m'  # Reset color to default

        colorized_data.append([f"{color_code}{old_text}{reset_color}", f"{color_code}{count}{reset_color}"])

    print(tabulate(colorized_data, headers=["Old Text", "Count"], tablefmt="fancy_grid"))


def main(in_xslx_file, col_search, col_replace, in_html_file, out_html_file):
    file_path = in_xslx_file  # Replace with your file path

    data_dict = excel_to_dict(file_path, col_search, col_replace)
    sorted_dict = sort_dict_by_key_length(data_dict)

    replaced_count = replace_text_in_html(in_html_file, out_html_file, sorted_dict)

    print_replacements_summary(replaced_count)


if __name__ == "__main__":
    os.system('cls' if os.name == 'nt' else 'clear')
    print(sys.getdefaultencoding())

    parser = argparse.ArgumentParser(description="Replace text in an HTML file based on XSLX table with translations")

    parser.add_argument("in_xslx_file", type=str, help="Dictionary XSLX file")
    parser.add_argument("col_search", type=str, help="Column to use for search")
    parser.add_argument("col_replace", type=str, help="Column to use for replace")
    parser.add_argument("in_html_file", type=str, help="Input HTML file")
    parser.add_argument("out_html_file", type=str, help="Output HTML file")

    args = parser.parse_args()

    in_xslx_file = args.in_xslx_file
    col_search = args.col_search
    col_replace = args.col_replace
    in_html_file = args.in_html_file
    out_html_file = args.out_html_file

    main(in_xslx_file, col_search, col_replace, in_html_file, out_html_file)
