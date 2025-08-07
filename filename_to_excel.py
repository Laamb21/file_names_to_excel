import os
import sys
import openpyxl
import argparse
from config import DEFAULT_COLUMN

def main():
    # Initialize argument parser
    parser = argparse.ArgumentParser(description="Read file names from a given directory, and copy them to an Excel file.")

    # Add arguments for agrument parser
    parser.add_argument(
        "directory",
        type=str,
        help="Path to the directory containing files."
    )
    parser.add_argument(
        "--output", 
        type=str,
        help='Path to the Excel file to write to. If not specified, a new Excel file named "file_names.xlsx" will be created in the directory.',
        default=None
    )
    parser.add_argument(
        "--sheet",
        type=str,
        help='Name of the sheet to write to. Default is the first (active) sheet.',
        default=None
    )
    parser.add_argument(
        "--column",
        type=str,
        help=f'Column letter to insert file names into. Default is "{DEFAULT_COLUMN}".',
        default=DEFAULT_COLUMN
    )
    parser.add_argument(
        "--append",
        action="store_true",
        help="If set, file names will be appended to the existing Excel file (if it exists). Otherwise, it will overwrite or create a new one."
    )