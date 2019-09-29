from unittest import TestCase
from ddt import ddt, idata, unpack
import os

from timeline_manager.get_command_line_parameters import get_command_line_parameters, process_command_line_parameters

test_command_line = (
    ('hello.xlsx', 'hello.xlsx', 'templates', 'output'),
    ('hello1.xlsx --templates template_folder', 'hello1.xlsx', 'template_folder', 'output'),
    ('hello2.xlsx --output output_folder', 'hello2.xlsx', 'templates', 'output_folder'),
    ('hello3.xlsx -t templates_folder3 --output output_folder3', 'hello3.xlsx', 'templates_folder3', 'output_folder3')
)

test_process_command_line = (
    (
        '/Users/thomasgaylardou/PycharmProjects/ppt-timeline3/testing/resources/command_line_test_folder/root_folder_01',
        'dummy_xl_file.xlsx',
        '/Users/thomasgaylardou/PycharmProjects/ppt-timeline3/testing/resources/command_line_test_folder/root_folder_01/dummy_xl_file.xlsx',
        '/Users/thomasgaylardou/PycharmProjects/ppt-timeline3/testing/resources/command_line_test_folder/root_folder_01/templates',
        '/Users/thomasgaylardou/PycharmProjects/ppt-timeline3/testing/resources/command_line_test_folder/root_folder_01/output'
    ),
    (
        '/Users/thomasgaylardou/PycharmProjects/ppt-timeline3/testing/resources/command_line_test_folder/root_folder_02',
        'xl_folder/dummy_xl_file2.xlsx -t template_folder --output output_folder',
        '/Users/thomasgaylardou/PycharmProjects/ppt-timeline3/testing/resources/command_line_test_folder/root_folder_02/xl_folder/dummy_xl_file2.xlsx',
        '/Users/thomasgaylardou/PycharmProjects/ppt-timeline3/testing/resources/command_line_test_folder/root_folder_02/template_folder',
        '/Users/thomasgaylardou/PycharmProjects/ppt-timeline3/testing/resources/command_line_test_folder/root_folder_02/output_folder'
    )
)


def generate_command_line_test_data():
    for command_line_data in test_command_line:
        command_line = command_line_data[0]
        xl = command_line_data[1]
        templates = command_line_data[2]
        output = command_line_data[3]

        yield command_line, 0, xl
        yield command_line, 1, templates
        yield command_line, 2, output


def generate_command_line_full_path_test_data():
    for command_line_data in test_process_command_line:
        current_working_dir = command_line_data[0]
        command_line = command_line_data[1]
        xl_full_path = command_line_data[2]
        template_full_path = command_line_data[3]
        output_full_path = command_line_data[4]

        yield current_working_dir, command_line, 0, xl_full_path
        yield current_working_dir, command_line, 1, template_full_path
        yield current_working_dir, command_line, 2, output_full_path


@ddt
class TestCommandLineProcessing(TestCase):
    @unpack
    @idata(generate_command_line_test_data())
    def test_command_line_01(self, command_line, parameter_index, exp_val):
        args = command_line.split()
        parameters = get_command_line_parameters(args)

        self.assertEqual(exp_val, parameters[parameter_index])

    @unpack
    @idata(generate_command_line_full_path_test_data())
    def test_process_command_line_01(self, cwd, command_line, parameter_index, exp_val):
        os.chdir(cwd)
        args = command_line.split()
        full_paths = process_command_line_parameters(args)

        self.assertEqual(exp_val, full_paths[parameter_index])
