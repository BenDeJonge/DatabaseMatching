# -*- coding: utf-8 -*-
"""
Transpose a table of trainings and followers.

Given are several .xslx files containing:
    n trainings with required followers.
These tables should be transposed to:
    a list of m jobtypes with required trainings.

    T1  T2  ... Tn
J1  T   V   ... T
J2  V   N   ... N
... ... ... ... ...
Jm  T   T   ... T

With the following abbreviations:
    Ji : the i-th job
    Ti : the i-th training
    T  : to be trained
    V  : to be viewed
    N  : not to be trained
"""

#-- I M P O R T S -------------------------------------------------------------
import openpyxl as xl
from re import split

#-- I N P U T S ---------------------------------------------------------------
xlsx_path = r'process_master_sheet.xlsx'
follower_col = 'B'
training_col = 'A'
viewing_col = 'C'

#-- C O N S T A N T S ---------------------------------------------------------
# Tracking all existing jobs and trainings.
all_training_types = ('train', 'watch')
jobs : dict = {}
all_trainings : set = set()

# Loading in workbook.
wb = xl.load_workbook(xlsx_path)
ws = wb.active
# Grabbing relevant columns.
xlsx_follower_col = [ cell.value for cell in ws[follower_col][1:] ]
xlsx_training_col = [ cell.value for cell in ws[training_col][1:] ]
xlsx_viewing_col  = [ cell.value for cell in ws[viewing_col][1:] ]

#-- M A I N   L O O P ---------------------------------------------------------
for training, trainees, watchers in zip(xlsx_training_col,
                                        xlsx_follower_col,
                                        xlsx_viewing_col):
    # Find and format all followers of a training.
    all_trainings.add(training)
    for followers, training_type in zip((trainees, watchers), 
                                        all_training_types):
        # Empty training.
        if not followers:
            continue
        # Followed training.
        followers = split(',|\n', followers)
        followers = [ f.strip() for f in followers ]
        for follower in followers:
            # Jobtype is not yet known.
            if not follower in jobs:
                jobs[follower] = {training_type : set() 
                                  for training_type in all_training_types}
            # Job does not yet follow training at higher intensity.
            if not training in jobs[follower][all_training_types[0]]:
                jobs[follower][training_type].add(training)
        
#-- W R I T I N G   O U T P U T -----------------------------------------------
# Creating output workbook.
wb = xl.Workbook()
ws = wb.active

# Writing headers in alphabetical order.
JOBCOL = 1
all_trainings = sorted(all_trainings)
ws['A1'] = 'Job'
for col, training in enumerate(all_trainings):
    _ = ws.cell(row=1, column=JOBCOL + col + 1, value = training)

# Writing data.
for i, (job, training_overview) in enumerate(jobs.items()):
    # Writing the job id.
    row = i + 2
    _ = ws.cell(row=row, column=JOBCOL, value=job)
    for training_type in all_training_types:
        # Getting all trainings for that job.
        trainings = training_overview[training_type]
        for training in trainings:
            # Writing the training type for each followed or watched training.
            _ = ws.cell(row=row, 
                        column=all_trainings.index(training) + 1 + JOBCOL,
                        value=training_type)       
# Outputting the file.
wb.save('test.xlsx')