class MilestoneElement:
    """
    Encapsulates all the information required for a single milestone within a timeline.  Includes each of the plotable
    elements (e.g. milestone shape, textbox etc)
    """
    def __init__(self, milestone_text, milestone_date, milestone_track_override=None, textbox_track_override=None):
        self.milestone_text = milestone_text
        self.milestone_date = milestone_date
        self.milestone_track_override = milestone_track_override
        self.textbox_track_override = textbox_track_override

