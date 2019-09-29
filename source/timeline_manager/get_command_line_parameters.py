import argparse
import os

def get_command_line_parameters(arguments=None):
    """

    :param args:
    :return: Return tuple of xxx elements, which contain the following:
             - Relative pathname of Excel file which drives the timeline to current working directory
             - Relative pathname of PowerPoint template files (relative to CWD)
             - Relative pathname where to place timeline PowerPoint files (relative to CWD)
    """
    parser = argparse.ArgumentParser()

    parser.add_argument("xl_file", help="Relative path of Excel file which has data for each timeline.")
    parser.add_argument("-t", "--templates",
                        default='templates',
                        help="Relative path of folder containing PowerPoint template files.")
    parser.add_argument("-o", "--output",
                        default='output',
                        help="Relative path of folder where the generated timelines will be placed.")

    if arguments is None:
        args = parser.parse_args()
    else:
        args = parser.parse_args(arguments)  # Only for testing.

    print(args)

    xl_path_param = args.xl_file
    template_folder_param = args.templates
    output_folder_param = args.output

    return xl_path_param, template_folder_param, output_folder_param


def check_relative_folder_path(path, folder=False):
    """
    Checks that a supplied string represents an existing folder as a relative path to the current working directory.
    return the full path of the folder if it is valid, raise an exception otherwise.

    :param path:
    :param folder:
    :return:
    """
    if not os.path.exists(path):
        raise ValueError(f'Excel file {path} does not exist')

    if folder is True:
        if not os.path.isdir(path):
            raise ValueError(f'{path} is not a folder')

    # Calculate full path for Excel file and then take folder as default folder for other parameters
    return os.path.abspath(path)


def process_command_line_parameters(*args):
    """
    Based on what the user has entered, calculate the full path for the Excel file, the template files and the output
    timeline files.

    Check that what is entered relates to real folders.

    ======================================================
    NOTE On relative paths, current working directory etc.
    ======================================================

    This has always confused me so this should help!

    1. When a python script executes, os.path.get_cwd() provides the folder which the user was at when the script was
       invoked.  NOT the folder where the script lives.

    2. Methods such as os.path.isdir() will accept a relative folder and will treat it as relative to the CWD.

    To avoid any ambiguity we will calculate the full path for each of the three folders and use those.


    :return:
    """
    xl_path_param, template_folder_param, output_folder_param = get_command_line_parameters(*args)

    xl_full_path = check_relative_folder_path(xl_path_param)
    template_full_path = check_relative_folder_path(template_folder_param, folder=True)
    output_full_path = check_relative_folder_path(output_folder_param, folder=True)

    return xl_full_path, template_full_path, output_full_path
