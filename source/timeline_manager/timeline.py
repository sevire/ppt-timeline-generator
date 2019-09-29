class Timeline:
    """
    Represents all the information required to generate a timeline in PowerPoint.
    """
    def __init__(self,
                 timeline_name,
                 parameter_data,
                 milestone_data):
        """
        Contains data for each milestone to be plotted.
        """
        self.name = timeline_name
        self.parameter_data = parameter_data
        self.milestone_data = milestone_data
