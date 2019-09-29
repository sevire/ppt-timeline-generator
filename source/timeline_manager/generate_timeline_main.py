import logging
import time
from source.timeline_manager.get_command_line_parameters import process_command_line_parameters
from source.timeline_manager.timeline_manager import TimelineManager

logger = logging.getLogger()

driver_data={
    'timeline_sheet_prefix': 'CPT'
}


def main():
    # ts = time.gmtime()
    # time_string = time.strftime("%Y_%m_%d__%H_%M_%S", ts)
    logging.basicConfig(filename='generate_timeline.log',
                        filemode='w',
                        format='%(levelname)s: %(asctime)s %(message)s',
                        level=logging.DEBUG
                        )

    wb_full_path, template_dir_fullpath, output_dir_full_path = process_command_line_parameters()

    logger.info('EPA Timeline Generator starting...\n\n')
    logger.info(f'Excel file is:      {wb_full_path}')
    logger.info(f'Template folder is: {template_dir_fullpath}')
    logger.info(f'Output folder is:   {output_dir_full_path}\n')

    timelines = TimelineManager(
        wb_full_path,
        template_dir_fullpath,
        output_dir_full_path,
        driver_data['timeline_sheet_prefix']
        )

    timelines.generate_timelines()


if __name__ == '__main__':
    main()




