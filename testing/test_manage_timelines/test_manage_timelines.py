from unittest import TestCase
import datetime
from ddt import ddt, idata, unpack
from pptx.util import Cm

from source.timeline_manager.timeline_manager import TimelineManager

test_xl_file_01 = '/Users/thomasgaylardou/Documents/EPA-timeline/testing/timeline_testfile_01.xlsx'

excel_test_data = (
    (
        'TEST-TIMELINE-01',
        {
            'start_date': datetime.datetime(2019, 4, 30),
            'end_date': datetime.datetime(2023, 12, 31),
            'include_ms_num_in_ms': False,
            'include_ms_num_in_text': False,
            'milestone_left': Cm(0.8),
            'milestone_right': Cm(34.81),
            'centre_vertical_position': Cm(15),
            'num_text_tracks': 6
        }
    ),
)


def test_data_gen():
    for timeline in excel_test_data:
        name = timeline[0]
        for key, val in timeline[1].items():
            yield name, key, val


@ddt
class TestManageTimelines(TestCase):
    def setUp(self) -> None:
        self.manager = TimelineManager(test_xl_file_01, timeline_sheet_prefix='TEST')

    def test_read_excel_data_01(self):
        timeline_list = self.manager.list_timelines()
        expected_list = [entry[0] for entry in excel_test_data]

        self.assertEqual(expected_list, timeline_list)

    @unpack
    @idata(test_data_gen())
    def test_read_excel_data_02(self, timeline_name, key, value):
        timeline = self.manager.get_timeline(timeline_name)

        self.assertIn(key, timeline.parameter_data)
        self.assertEqual(value, timeline.parameter_data[key])

    def test_plot_timeline_01(self):
        self.manager.generate_timelines()

        pass
