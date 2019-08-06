import pandas as pd


def read_milestone_data(workbook_name, sheet_name):

    milestones = pd.read_excel(workbook_name, sheet_name)
    milestones.set_index('Milestone Number', inplace=True)

    return milestones
