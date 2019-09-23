import logging

from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE

logger = logging.getLogger()


class ShapeWrapper:
    """
    Used as a wrapper to a PowerPoint shape to help with applying similar formatting to other shapes.  Note this
    is required due to a gap in the functionality of python-pptx which doesn't allow a shape to be copied (I think).

    """

    def __init__(self, template_shape):
        """
        Takes a supplied shape as a template.  This shape will be used then to create new shapes where all attributes
        are the same, other than those which are overridden (such as position).

        :param template_shape:
        """

        self.shape = template_shape

    def create_new_shape(self, shapes_object, left=None, top=None, width=None, height=None):
        """
        Add new shape to shapes object which is identical to the original template shape, except where allowed
        attributes have been overridden.

        :param shapes_object: PowerPoint shapes object to which the new shape will be added.
        :param left:   If specified then override left from template
        :param top:    If specified then override top from template
        :param width:  If specified then override width from template
        :param height: If specified then override height from template
        :return:
        """

        if self.shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            auto_shape_type = self.shape.auto_shape_type
            shape_dimensions = (
                self.shape.left if left is None else left,
                self.shape.top if top is None else top,
                self.shape.width if width is None else width,
                self.shape.height if height is None else height
            )
            new_shape = shapes_object.add_shape(auto_shape_type, *shape_dimensions)

            self.set_fill(new_shape)
            self.set_line(new_shape)
        else:
            print('Shape types other than auto shapes not currently supported')

    def set_fill(self, shape):
        if self.shape.fill.type == MSO_FILL_TYPE.SOLID:
            new_fill = shape.fill
            new_fill.solid()
            new_fore_colour = new_fill.fore_color
            new_fore_colour.rgb = self.shape.fill.fore_color.rgb
        else:
            print('Fill type of {} not currently supported'.format(self.shape.fill.type))

    def get_fill(self):
        fill = self.shape.fill
        if fill.type == MSO_FILL_TYPE.SOLID:
            foreground_colour = fill.fore_color
            rgb = foreground_colour.rgb
            return rgb
        else:
            return '(unknown)'

    def set_line(self, shape):
        line = shape.line
        line.type = 1
        line.color.rgb = self.shape.line.color.rgb
        line.width = self.shape.line.width

    def __str__(self):
        formatted_string = \
            'Shape type is:     {}\n'.format(self.shape.shape_type, ) + \
            'Autoshape type is: {}\n'.format(self.shape.auto_shape_type) + \
            'Fill colour:       {!s}'.format(self.get_fill())

        return formatted_string

