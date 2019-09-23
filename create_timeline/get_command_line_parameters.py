import argparse
from pathlib import Path


def get_command_line_parameters():
    """

    :param args:
    :return: Return tuple of xxx elements, which contain the following:
             - Relative pathname of Excel file which drives the timeline to current working directory
             - Relative pathname of PowerPoint template files (relative to CWD)
             - Relative pathname where to place timeline PowerPoint files (relative to CWD)
    """
    parser = argparse.ArgumentParser()

    parser.add_argument("xl_param_file", help="Relative path of file to use with timeline data.")
    parser.add_argument("ppt_template_folder", help="Relative path of folder containing PowerPoint template files")
    parser.add_argument("ppt_output_folder", help="Relative path of folder containing PowerPoint timeline files")

    args = parser.parse_args()

    print(args)

    x = args.xl_param_file
    y = args.ppt_template_folder
    z = args.ppt_output_folder

    return x, y, z
