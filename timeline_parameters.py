from collections import namedtuple

import pandas as pd
from pptx.util import Cm

TrackInfo = namedtuple('TrackInfo', 'separation centre start')

# Set parameters which have literal values
timeline_param = {
    'CPT-HAHAP-01': {
        'start_date': pd.to_datetime('2017-07-01'),
        'end_date': pd.to_datetime('2020-08-31'),
        'include_ms_num_in_ms': False,
        'include_ms_num_in_text': False,
        'milestone_left': Cm(0.75),
        'milestone_right': Cm(34.1 + 0.71),
        'centre_vertical_position': Cm(15),
        'num_text_tracks': 5,
        'TrackInfo': namedtuple('TrackInfo', 'separation centre start'),
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
    },
    'CPT-HAHAP-02': {
        'start_date': pd.to_datetime('2019-07-01'),
        'end_date': pd.to_datetime('2022-08-31'),
        'include_ms_num_in_ms': False,
        'include_ms_num_in_text': False,
        'milestone_left': Cm(0.75),
        'milestone_right': Cm(34.1 + 0.71),
        'centre_vertical_position': Cm(15),
        'num_text_tracks': 5,
        'TrackInfo': namedtuple('TrackInfo', 'separation centre start'),
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
}

# Now parameters with calculated values
for timeline_record in timeline_param.values():
    timeline_record['milestone_total_days'] = (timeline_record['end_date'] - timeline_record['start_date']).days + 1
    timeline_record['milestone_x_range'] = timeline_record['milestone_right'] - timeline_record['milestone_left']
    timeline_record['milestone_track_data'] = TrackInfo(Cm(0.8), timeline_record['centre_vertical_position'], Cm(0.8))
    timeline_record['textbox_track_data'] = TrackInfo(Cm(1.2), timeline_record['centre_vertical_position'], Cm(2.95))
