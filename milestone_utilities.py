from datetime import datetime as dt
from pptx_utilities import create_line, create_shape_from_template
import logging
from timeline_parameters import timeline_param as tlp

logger = logging.getLogger()


def calculate_next_track(milestone_id, day_num, orientation):
    """
    Works out which part (third probably) of chart this day appears on then retrieves track value to use and
    updates for next call.

    :param milestone_id:
    :param day_num:
    :param orientation:
    :return:
    """
    area_num = chart_area_num(milestone_id, day_num)
    current_track = tlp[milestone_id]['textbox_track_position_data'][area_num]['track'][orientation]
    tlp[milestone_id]['textbox_track_position_data'][area_num]['track'][orientation] = \
        ((current_track - 1 + tlp[milestone_id]['textbox_track_position_data'][area_num]['direction']) %
         tlp[milestone_id]['num_text_tracks']) + 1

    return current_track * orientation


# ### Utility functions for working out where to place things
def day_num_as_proportion(milestone_id, day): return (day - 1) / (tlp[milestone_id]['milestone_total_days'] - 1)


def sign(x): return (1, -1)[x <= 0]  # Implement sign function as isn't standard


def to_day_num(milestone_id, date):
    record = tlp[milestone_id]
    start_date = record['start_date']
    return (date - start_date).days + 1


def parse_track_override(override_value):
    """
    The spreadsheet can contain a track number to override the calculated one for either milestone or text box.  This
    can help with fine tuning positioning for overlapping elements or just to control which elements line up with
    which others.

    If the override cell is left blank in a particular row, then it will come through as Nan.  We need to convert to
    None.  Also we can check that the track number isn't a stupid value.

    :param override_value:
    :return: Integer track number or None
    """
    try:
        track_number = int(override_value)
        return track_number
    except ValueError:
        return None
    except Exception:
        print('Unexpected exception converting track_override {}'.format(override_value))
        return None


def chart_area_num(milestone_id, day_num):
    """
    Calculates which part of the chart this day will appear in.  One of:
        1: Left
        2: Middle
        3: Right

    :param milestone_id:
    :param day_num: Which day within the date range for the chart this is
    :return: 1, 2, or 3.
    """
    return int(day_num_as_proportion(milestone_id, day_num) * 3 + 1)


def calculate_text_box_shift(milestone_id, day_num, text_box_width, milestone_width):
    """
    Calculates the correction to make to a text box's left position to take account of where it is on the slide.

    If it is in the left hand third of the page, then it will be shifted right in order to avoid overlapping with
    the left edge.  The left edge of the text box will line up with the left edge of the milestone shape

    If it is in the right third of the page then it will be shifted left to avoid overlapping with
    right edge.

    If it is in the centre third then it will be left as centred.

    :param milestone_id:
    :param day_num:         Used to calculate which third of the chart the text box is
    :param text_box_width:  Used to calculate how far to move the text box
    :param milestone_width: Used to help line up text box with edge of milestone shape
    :return:
    """
    area_num = chart_area_num(milestone_id, day_num)
    if area_num == 1:
        # Left hand side of chart
        return (text_box_width - milestone_width) // 2
    elif area_num == 3:
        # Right hand side of chart
        return (-text_box_width + milestone_width) // 2
    else:
        return 0


def calculate_offset_for_text_track(milestone_id, orientation, day_num, track_number=None):
    """
    Similar to calculate_offset_for_ms_track, although there is no level to base the track number on.  Instead
    we will choose the track using two factors:

    - The orientation (+1 or -1) will be passed in and the selected track number will conform to that so that the
      text box is on the same side of the centre line as the milestone.

    - A record will be kept of the most recent track number for each orientation and at each call the next number
      will be selected (mod the number of tracks per side of the centre line)

    :param milestone_id:
    :param orientation:
    :param day_num:
    :param track_number:
    :return:
    """
    if track_number is not None:
        track = track_number
    else:
        track = calculate_next_track(milestone_id, day_num, orientation)

    position = calculate_track_location(track, *tlp[milestone_id]['textbox_track_data'])

    # Calculate the (relative) amount off the centre line we need to place the centre of the text box
    # vertical_offset_rel = text_box_track_start_offset + (track - 1) * offset_per_text_track

    # Now calculate the absolute offset based on where the centre line is and which orientation we are plotting
    # vertical_position = centre_vertical_position + orientation * vertical_offset_rel

    return position


def calculate_track_location(track_number, track_separation, centre_position, track_start_offset=0):
    """
    General purpose function to calculate the vertical position of where the centre of an element needs to be to line
    up with the specified track.  Will be used for both milestone and textbox tracks (and others if required later).

    The way tracks will work is that different element types (e.g. milestones or associated text boxes) will be aligned
    to pre-determined horizontal 'tracks'.  Tracks can be above or below a specified centre line.

    If the tracks include the centre line then the centre line has track number zero.  For other cases a track of zero
    has no meaning and is invalid.

    If the tracks don't include the centre line then the first track will start at a specified distance from the centre
    line and all other tracks positioned relative to it.

    :param track_number:       Signed integer denoting the number of the track at which the element is to be placed.
                               - 0  is the centre track but this is only meaningful when the set of tracks includes the
                                    centre (i.e if the track_start_offset is zero).
                               - +n is the nth track below the centre line (using the PowerPoint plotting convention that
                                    the y coordinate starts at the top and increases as it goes down the page.
                               - -n is the nth track above the centre line.

    :param track_separation:   The vertical distance between tracks

    :param centre_position:    Vertical position (supplied as pptx.util.Cm probably) of centre line of tracks.  Used as
                               anchor point to calculate the offsets for tracks from.

    :param track_start_offset: If the set of tracks doesn't include the centre line then this value specifies how far
                               from the centre line the first track should be.

    :return:                   Distance from the top of the slide where the track is located.
    """
    orientation = sign(track_number)
    abs_track_num = abs(track_number)
    abs_track_offset = track_start_offset + (abs_track_num-1) * track_separation
    position = centre_position + orientation * abs_track_offset

    return position


def get_plot_data(milestone_id, day_num, milestone_track, orientation, textbox_track_override):
    """
    Calculates the position of each of the elements of the milestone, which are:
    - Milestone indicator (circle)
    - Milestone description (rectangle)
    - Connecting line (line)

    The arrangement needs to be as follows:

    The track number is a positive or negative integer which describes how far away the milestone element needs
    to be from the centre line, and whether it is above or below it.

    The milestone indicator will be positioned around the centre line at a vertical distance determined by the
    track number allocated to the milestone and a horizontal distance which places it precisely on the right date
    for the milestone.

    The text box will be placed in the same configuration as the milestone indicator, but further away from the
    centre line.

    The connecting line will run from the edge of the milestone indicator in the direction of its corresponding
    text box, to the nearest edge of the text box.

    What will be returned will be three tuples, which indicate where the CENTRE of the corresponding element
    needs to be.  The caller will then need to adjust depending upon the width and height of each element.

    The tuples are structured as follows:

    - milestone_indicator_data  = (left, top)
    - milestone_description_box = (left, top)
    - connecting_line           = (left_1, top_1, left_2, top_2)

    :param milestone_id:
    :param day_num:
    :param milestone_track:
    :param orientation:
    :param textbox_track_override:
    :return:
    """
    # Calculate x position first as this will be used in plotting of all elements.
    x_rel = day_num_as_proportion(milestone_id, day_num) * tlp[milestone_id]['milestone_x_range']
    x_abs = int(tlp[milestone_id]['milestone_left'] + x_rel)

    # Calculate data to drive plotting of each element

    # milestone_vertical_position = calculate_offset_for_ms_track(milestone_track)
    milestone_vertical_position = calculate_track_location(milestone_track, *tlp[milestone_id]['milestone_track_data'])
    milestone_element_data = (x_abs, milestone_vertical_position)

    # --- Second the text box
    textbox_vertical_position = calculate_offset_for_text_track(milestone_id, orientation,
                                                                day_num, textbox_track_override)
    textbox_element_data = (x_abs, textbox_vertical_position)

    return milestone_element_data, textbox_element_data  # Not including connector line yet


def create_milestone_shapes(timeline_name, milestones, shapes_object, templates):
    """
    Iterate through the milestone data and calculate the information required to create new shapes.  Store it until
    all the data has been collected, then go back through the data three times to create the shapes in the order of:
    - Connectors
    - Milestones
    - Text boxes

    As new shapes are placed in front of older shapes, this will ensure that the connectors travel behind the shapes.

    :param timeline_name:
    :param milestones:
    :param shapes_object:
    :param templates:
    :return:
    """

    # track_orientation = 1

    # Create the structure to store all the data to plot the points, but plot later to control what appears on top
    # of what.
    shape_data = {
        'connector': {
            'func': create_line,
            'parameters': []
        },
        'milestone': {
            'func': create_shape_from_template,
            'parameters': []
        },
        'textbox': {
            'func': create_shape_from_template,
            'parameters': []
        }
    }
    for milestone_number, milestone_data in milestones.iterrows():
        milestone_name = milestone_data['Milestone Name']
        milestone_date = milestone_data['Date']
        milestone_level = milestone_data['Milestone Level']  # Normalises track to centre on zero
        milestone_level_text_label = 'TEXT {}'.format(milestone_level)
        milestone_track_override = parse_track_override(milestone_data['Milestone Track Override'])
        textbox_track_override = parse_track_override(milestone_data['Textbox Track Override'])

        # Plot milestone shape at right position
        day_num = to_day_num(timeline_name, milestone_date)

        # Calculate milestone position and whether above or below centre line
        milestone_track_number = milestone_track_override if milestone_track_override is not None \
            else (milestone_level - 1) * tlp[timeline_name]['milestone_track_orientation'][milestone_level]

        milestone_data, textbox_data = \
            get_plot_data(timeline_name, day_num, milestone_track_number,
                          tlp[timeline_name]['milestone_track_orientation'][milestone_level], textbox_track_override)

        # Create and position milestone marker shape
        ms_x, ms_y = milestone_data

        milestone_width = templates[str(milestone_level)]['data']['width']
        milestone_height = templates[str(milestone_level)]['data']['height']

        # Extract fill color of milestone as we will re-use this for connector (for the moment)
        milestone_colour = templates[str(milestone_level)]['data']['fill_colour_rgb']

        ms_left = ms_x - milestone_width // 2
        ms_top = ms_y - milestone_height // 2

        milestone_text = str(milestone_number) if tlp[timeline_name]['include_ms_num_in_text'] else ''

        # Add milestone data to data structure for plotting later
        shape_data['milestone']['parameters'].append(
            (shapes_object, ms_left, ms_top, templates[str(milestone_level)]['data'], milestone_text)
        )

        # Create and position text box shape
        text_x, text_y = textbox_data

        textbox_width = templates[milestone_level_text_label]['data']['width']
        textbox_height = templates[milestone_level_text_label]['data']['height']

        tx_left = text_x - textbox_width // 2 + calculate_text_box_shift(timeline_name, day_num,
                                                                         textbox_width, milestone_width)
        tx_top = text_y - textbox_height // 2

        display_date = dt.strftime(milestone_date, '%d-%b-%Y')
        if tlp[timeline_name]['include_ms_num_in_text'] is True:
            textbox_text = '({}) [{}]\n{}'.format(milestone_number, display_date, milestone_name)
        else:
            textbox_text = '[{}]\n{}'.format(display_date, milestone_name)

        # Add text box data to data structure for plotting later
        shape_data['textbox']['parameters'].append(
            (shapes_object, tx_left, tx_top, templates['TEXT {}'.format(milestone_level)]['data'], textbox_text)
        )

        # Create and position connecting line.  Note it won't actually connect to anything but will be placed correctly
        line_x = ms_x

        # Plot connector from middle of shapes (for now) to avoid having to re-factor to properly process tracks
        # ToDo: Refactor this so that tracks of ms and tx are used to calculate connector line

        line_ms_y = ms_y
        line_text_y = text_y

        # line_ms_y = ms_y + tlp[timeline_name]['milestone_track_orientation'][milestone_level] * milestone_height // 2
        # line_text_y = text_y - tlp[timeline_name]['milestone_track_orientation'][milestone_level] * textbox_height // 2

        # Add connector data to data structure for plotting later
        shape_data['connector']['parameters'].append((shapes_object, line_x, line_ms_y, line_text_y, milestone_colour))

        # Alternate between positive and negative track numbers to minimise overlap of objects
        tlp[timeline_name]['milestone_track_orientation'][milestone_level] = \
            -tlp[timeline_name]['milestone_track_orientation'][milestone_level]

    # Now plot the data in the order of connectors (bottom), milestones and text boxes
    for shape_type in shape_data:
        func = shape_data[shape_type]['func']
        for param in shape_data[shape_type]['parameters']:
            func(*param)
