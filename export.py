import pandas as pd
from bs4 import BeautifulSoup, Comment
import re
from tabulate import tabulate
import argparse
import os


def extract_text_out_html(input_file):
    with open(input_file, 'r') as file:
        html_content = file.read()

    export_dict = {}
    soup = BeautifulSoup(html_content, 'html.parser')

    for tag in soup.find_all(string=True):
        if tag.parent.name not in ['style', 'script', 'head', 'title', 'meta', '[document]']:
            if not isinstance(tag, Comment):
                export_dict[tag.text] = tag.text

    return list(export_dict.keys())


def save_to_excel(data, output_file):
    df = pd.DataFrame({'Text': data})
    df.to_excel(output_file, index=False)


def main(in_html_file, out_xslx_file):
    data = extract_text_out_html(in_html_file)
    #print(data)
    print(tabulate({"text": data}, headers=["Text"], tablefmt="fancy_grid"))
    save_to_excel(data, out_xslx_file)


if __name__ == "__main__":
    os.system('cls' if os.name == 'nt' else 'clear')

    parser = argparse.ArgumentParser(description="Replace text in an HTML file based on XSLX table with translations")

    parser.add_argument("in_html_file", type=str, help="Input HTML file")
    parser.add_argument("out_xslx_file", type=str, help="Dictionary XSLX file")

    args = parser.parse_args()

    in_html_file = args.in_html_file
    out_xslx_file = args.out_xslx_file

    main(in_html_file, out_xslx_file)
