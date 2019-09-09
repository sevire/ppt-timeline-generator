import pandas as pd
from collections import namedtuple

from pptx.util import Cm

timeline_excel_read_data = {
    'parameters_table_start_row': 2,
    'parameters_table_num_rows': 8,
    'parameters_table_column_range': 'B:D',
    'timeline_table_start_row': 13,
    'valid_sheet_prefix': 'CPT'
}

TrackInfo = namedtuple('TrackInfo', 'separation centre start')


def get_list_of_timeline_sheets(workbook_full_pathname: str):
    xl = pd.ExcelFile(workbook_full_pathname)
    sheet_names = xl.sheet_names
    valid_sheets = filter(lambda x: x[:3] == timeline_excel_read_data['valid_sheet_prefix'], sheet_names)
    return xl, valid_sheets


def extract_parameter_data(timeline_name, param_dict, pandas_object):
    param_dict[timeline_name] = {}

    param_dict[timeline_name]['start_date'] = pandas_object.loc['start_date', 'Value']
    param_dict[timeline_name]['end_date'] = pandas_object.loc['end_date', 'Value']
    param_dict[timeline_name]['include_ms_num_in_ms'] = pandas_object.loc['include_ms_num_in_ms', 'Value']
    param_dict[timeline_name]['include_ms_num_in_text'] = pandas_object.loc['include_ms_num_in_text', 'Value']
    param_dict[timeline_name]['milestone_left'] = Cm(pandas_object.loc['milestone_left', 'Value'])
    param_dict[timeline_name]['milestone_right'] = Cm(pandas_object.loc['milestone_right', 'Value'])
    param_dict[timeline_name]['centre_vertical_position'] = Cm(pandas_object.loc['centre_vertical_position', 'Value'])
    param_dict[timeline_name]['num_text_tracks'] = pandas_object.loc['num_text_tracks', 'Value']

    param_dict[timeline_name]['milestone_track_orientation'] = {
        1: 1,
        2: 1,
        3: 1
    }

    param_dict[timeline_name]['textbox_track_position_data'] = {
        1: {
            'name': 'left',
            'track': {1: 5, -1: 5},
            'direction': -1
        },
        2: {
            'name': 'middle',
            'track': {1: 5, -1: 5},
            'direction': -1
        },
        3: {
            'name': 'right',
            'track': {1: 1, -1: 1},
            'direction': 1
        }
    }

    param_dict[timeline_name]['milestone_total_days'] = \
        (param_dict[timeline_name]['end_date'] - param_dict[timeline_name]['start_date']).days + 1
    param_dict[timeline_name]['milestone_x_range'] = \
        param_dict[timeline_name]['milestone_right'] - param_dict[timeline_name]['milestone_left']
    param_dict[timeline_name]['milestone_track_data'] = \
        TrackInfo(Cm(0.8), param_dict[timeline_name]['centre_vertical_position'], Cm(0.8))
    param_dict[timeline_name]['textbox_track_data'] = \
        TrackInfo(Cm(1.2), param_dict[timeline_name]['centre_vertical_position'], Cm(2.95))


def read_parameter_data(xl, sheet_name):
    parameters = xl.parse(sheet_name,
                          skiprows=timeline_excel_read_data['parameters_table_start_row']-1,
                          nrows=timeline_excel_read_data['parameters_table_num_rows'],
                          usecols=timeline_excel_read_data['parameters_table_column_range'])
    parameters.set_index('Parameter Name', drop=False, inplace=True)
    return parameters


def read_milestone_data(xl, sheet_name):
    milestones = xl.parse(sheet_name, skiprows=timeline_excel_read_data['timeline_table_start_row']-1)
    milestones.set_index('Milestone Number', inplace=True)

    return milestones
