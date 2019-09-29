from copy import deepcopy

from pptx.enum.shapes import MSO_CONNECTOR_TYPE


class TimelineShape:
    """
    Represents the data required to manage a PowerPoint shape object which will be used to plot an element of the
    timeline within a PowerPoint slide.

    The ppt_object needs to be an autoshape (at the time of writing) otherwise the logic gets too complex for the
    application. (i.e. it can't be a custom shape with user added vertices etc.).

    Designed to be a superclass of specific shapes
    """
    def __init__(self, ppt_object):
        """

        :param ppt_object:
        :param shapes_object: None if ppt_object is passed in.  Populated if shape is to be created from scratch.
        """
        self.ppt_object = ppt_object

    def clone_shape(self, shape):
        """
        Copies any common elements.  Any specific elements will need to be implemented by sub-class.

        :param shape:
        :return:
        """
        pass


class MilestoneShape(TimelineShape):
    pass


class TextboxShape(TimelineShape):
    pass


class ConnectorShape(TimelineShape):
    def __init__(self, ppt_shape):
        super().__init__(ppt_shape)

    pass

    @classmethod
    def from_ppt_shape(cls, shapes_object, ppt_shape):
        pass

    @classmethod
    def from_parameters(cls, shapes_object, x, y_ms, y_text, rgb_colour):

        line = shapes_object.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT, x, y_ms, x, y_text)
        line.line.color.rgb = rgb_colour

        return ConnectorShape(line)
