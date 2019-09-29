class Milestone:
    """
    Represents key information about a milestone within a timeline which will allow it to be plotted in a PowerPoint
    representation of the timeline.
    """
    def __init__(self,
                 milestone_number,
                 milestone_text,
                 milestone_date,
                 milestone_level,
                 milestone_track_number,
                 textbox_track_number,
                 plotting_parameters):

        """
        Define key properties of a milestone so that it can be plotted.

        :param milestone_number:
        :param milestone_text:
        :param milestone_date:
        :param milestone_level:
        :param milestone_track_number:
        :param textbox_track_number:
        :param plotting_parameters:
        """
        self.milestone_number = milestone_number
        self.milestone_text = milestone_text
        self.milestone_date = milestone_date
        self.milestone_level = milestone_level
        self.milestone_track_number = milestone_track_number
        self.textbox_track_number = textbox_track_number

