# -*- coding: utf-8 -*-
"""
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
from openpyxl import Workbook
from re import split

#-- C O N S T A N T S ---------------------------------------------------------
training_types = ('training', 'viewing')
dummy_job : dict = {training_type : [] for training_type in training_types}
separator = ','

# 
jobs : dict = {}        # Read from .xlsx
trainings : set = set() # Read from .xlsx. How to keep wanted order?

# To be read as pd.DataFrame.
xlsx_follower_col = []
xlsx_training_col = []
xlsx_viewing_col  = []

#-- M A I N   L O O P ---------------------------------------------------------
for training, training_type in zip((xlsx_training_col, xlsx_viewing_col),
                                   training_types):
    # Find and format all followers of a training.
    followers = split(',|\n', xlsx_follower_col[training])
    followers = [ f.strip() for f in followers ]
    # Add training follower to job list. Trainings are classified.
    for follower in followers:
        if not follower in jobs:
            jobs[follower] = dummy_job
        jobs[follower][training_type].append(training)
        
#-- W R I T I N G   O U T P U T -----------------------------------------------
wb = Workbook()
ws = wb.active
ws.title('Trainings')

