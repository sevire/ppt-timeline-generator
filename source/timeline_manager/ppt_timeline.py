import os

from pptx import Presentation

from source.timeline_manager.timeline import Timeline
import logging

from timeline_manager.ppt_timeline_file import PptTimelineFile

logger = logging.getLogger(__name__)


class PptTimeline:
    def __init__(self, timeline: Timeline, template_dir, output_dir):
        self.timeline = timeline
        self.template_dir = template_dir
        self.output_dir = output_dir

    def generate_timeline(self):
        template_filename = os.path.join(self.template_dir, self.timeline.name + '.pptx')
        generated_filename = os.path.join(self.output_dir, self.timeline.name + '-out.pptx')

        logger.info(f'Template file for {self.timeline.name} is {template_filename}')
        logger.info(f'Output file for {self.timeline.name} is {generated_filename}')

        ppt_timeline = PptTimelineFile(template_filename)
        ppt_timeline.generate_plot_data()
        ppt_timeline.plot_timeline()
        ppt_timeline.save_timeline(generated_filename)


        # Expect to find template objects in first slide - don't look anywhere else
        template_slide = slides[0]
        template_shapes = template_slide.shapes
        template_data = self.extract_template_data(template_shapes)

        # We are adding shapes to the second slide.
        milestone_slide = slides[1]
        milestone_shapes = milestone_slide.shapes
        self.create_milestone_shapes(self.timeline.name,
                                     self.timeline.milestone_data,
                                     milestone_shapes,
                                     template_data,
                                     self.timeline.parameter_data)

        prs.save(presentation_name_out)







