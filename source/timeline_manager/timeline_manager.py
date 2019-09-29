import glob
import os
from collections import namedtuple
import pandas as pd
from pptx.util import Cm
from source.timeline_manager.timeline import Timeline
from source.timeline_manager.ppt_timeline_file import PptTimelineFile
import logging

logger = logging.getLogger(__name__)

TrackInfo = namedtuple('TrackInfo', 'separation centre start')

timeline_excel_read_data = {
    'parameters_table_start_row': 2,
    'parameters_table_num_rows': 8,
    'parameters_table_column_range': 'B:D',
    'timeline_table_start_row': 13,
    'valid_sheet_prefix': 'CPT',
    'boolean_true': 'Yes'
}


class TimelineManager:
    """
    Manages the information held in Excel which contains the milestone data and also the data which drives the plotting
    of the timelines (e.g. where the left and right hand edges of the timeline are on the slide).
    """

    def __init__(self,
                 excel_file_full_pathname,
                 template_folder_full_pathname='templates',
                 output_folder_full_pathname='output',
                 timeline_sheet_prefix='CPT'):
        """
        Defines where the key data which is needed to create the timeline.

        To simplify things, the only parameter which is mandatory is the pathname of the Excel file, relative to where
        the script is run from (i.e the CWD of the terminal session at the point the script is run, not the location
        of the script).

        By default, the template files and the output files will be placed in named folders below the folder where the
        Excel file is, but these can be overridden.

        :param excel_file_full_pathname:
        :param template_folder_full_pathname:
        :param output_folder_full_pathname:
        """
        self.excel_file_full_pathname = excel_file_full_pathname
        self.template_folder_full_pathname = template_folder_full_pathname
        self.output_folder_full_pathname = output_folder_full_pathname

        self.timeline_sheet_prefix = timeline_sheet_prefix

        # Dummy initialise for properties which will be properly initialised outside here.
        self.timelines = {}
        self.xl_data = None

        self.read_excel_data()

    def read_excel_data(self):
        self.xl_data = pd.ExcelFile(self.excel_file_full_pathname)

        sheet_names = self.xl_data.sheet_names

        for name in sheet_names:
            parameter_data = self.read_parameter_data(name)
            milestone_data = self.read_timeline_data(name)
            timeline = Timeline(name, parameter_data, milestone_data)
            self.timelines[name] = timeline

    def list_timelines(self):
        return list(self.timelines.keys())

    def get_timeline(self, timeline_name):
        if timeline_name in self.timelines:
            return self.timelines[timeline_name]
        else:
            return None

    def read_parameter_data(self, timeline_name):
        parameters = self.xl_data.parse(timeline_name,
                                        skiprows=timeline_excel_read_data['parameters_table_start_row'] - 1,
                                        nrows=timeline_excel_read_data['parameters_table_num_rows'],
                                        usecols=timeline_excel_read_data['parameters_table_column_range'])
        parameters.set_index('Parameter Name', drop=False, inplace=True)

        parameter_data = {
            'start_date': parameters.loc['start_date', 'Value'],
            'end_date': parameters.loc['end_date', 'Value'],
            'include_ms_num_in_ms': parameters.loc['include_ms_num_in_ms', 'Value'] ==
            timeline_excel_read_data['boolean_true'],
            'include_ms_num_in_text': parameters.loc['include_ms_num_in_text', 'Value'] ==
                                      timeline_excel_read_data['boolean_true'],
            'milestone_left': Cm(parameters.loc['milestone_left', 'Value']),
            'milestone_right': Cm(parameters.loc['milestone_right', 'Value']),
            'centre_vertical_position': Cm(
                parameters.loc['centre_vertical_position', 'Value']),
            'num_text_tracks': parameters.loc['num_text_tracks', 'Value'],
            'milestone_track_orientation': {
                1: 1,
                2: 1,
                3: 1
            },
            'textbox_track_position_data': {
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
        }

        # Add values which depend on other values so can't be initialised at the same time.
        parameter_data['milestone_total_days'] = (parameter_data['end_date'] - parameter_data['start_date']).days + 1
        parameter_data['milestone_x_range'] = parameter_data['milestone_right'] - parameter_data['milestone_left']
        parameter_data['milestone_track_data'] = TrackInfo(Cm(0.8), parameter_data['centre_vertical_position'], Cm(0.8))
        parameter_data['textbox_track_data'] = TrackInfo(Cm(1.2), parameter_data['centre_vertical_position'], Cm(2.95))

        return parameter_data

    def read_timeline_data(self, timeline_name):
        milestones = self.xl_data.parse(timeline_name,
                                        skiprows=timeline_excel_read_data['timeline_table_start_row'] - 1)
        milestones.set_index('Milestone Number', inplace=True)

        return milestones

    def generate_timelines(self):
        """
        Iterates through all the timelines capture from the Excel file, and generates the timeline for each one.

        :return: None
        """
        for timeline in self:
            input_file = os.path.join(self.template_folder_full_pathname, timeline + '.pptx')
            output_file = os.path.join(self.output_folder_full_pathname, timeline + '_out.pptx')
            ppt_timeline = PptTimelineFile(input_file, output_file, self.timelines[timeline])
            ppt_timeline.generate_timeline()

    def delete_old_timelines(self):
        path = os.path.join(self.output_folder_full_pathname, '*.out')
        file_list = glob.glob(path)
        for file_path in file_list:
            try:
                os.remove(file_path)
            except:
                logger.warning(f"Error while deleting old output file : {file_path}")

    def __len__(self):
        return len(self.list_timelines())

    def __iter__(self):
        return iter(self.timelines)
