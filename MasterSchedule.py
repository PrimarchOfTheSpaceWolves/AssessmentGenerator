import datetime
import urllib.request as ul
from bs4 import BeautifulSoup as soup
import numpy as np
import pandas as pd
import requests
import Banner as bn

'''
# TODO: Revamp using requests library and soup subparsing
def __split_columns(data, begin, to_remove):
    for t in to_remove:
        data = data.replace(t, "")            
    tokens = data.split(begin)      
    tokens = tokens[1:]      
    return tokens 
'''

def get_subject(year, semester, subj, 
                specific_course_nums=None,
                columns_to_keep=None,
                add_semester_year_cols=False,
                remove_empty_courses=True):
        
    # Do web query
    term_in = bn.get_term_code(year, semester)
    
    url="https://banner.sunypoly.edu/pls/prod/swssschd.P_ShowSchd"    
    data = {
        "term_in" : term_in,
        "disc_in" : subj
    }    
    session = requests.Session() 
    x = session.post(url, data = data)

    # Get all table data
    all_data_pd = bn.parse_banner_html_table_data(x.text, "Master Schedule")
    
    # If requested, only keep specific courses
    if specific_course_nums is not None:        
        all_data_pd = all_data_pd.loc[all_data_pd["Crs"].isin(specific_course_nums)]
        
    # If requested, remove empty courses
    if remove_empty_courses:
        all_data_pd = all_data_pd.loc[all_data_pd["ENL"] != "0"]           
    
    # If asked, only keep certain columns
    if columns_to_keep is not None:
        all_data_pd = all_data_pd[columns_to_keep]

    # Add semester and year columns?
    if add_semester_year_cols:
        all_data_pd.insert(0, "Year", year)
        all_data_pd.insert(1, "Semester", semester)
        
    # Reset index
    all_data_pd.reset_index(drop=True, inplace=True)  

    return all_data_pd        

  
# year_semester_list - list with form ["2022 Spring", "2021 Summer"]
# subjects - dictionary with form { "CS" : None,
#                                   "MAT" : ["115", "413"]}
#           (basically, subject with either None or list of specific courses)
def get_courses(year_semester_list,
                subjects,                 
                columns_to_keep=None):

    # Start with empty list of dataframes
    all_data_frames_list = []

    # For each desired semester...
    for year_sem in year_semester_list:
        # Get year and semester
        year, semester = year_sem.split()
        # For each subject...
        for subj in subjects:
            # Get desired courses
            specific_course_nums = subjects[subj]
            # Get dataframe 
            semester_dataframe = get_subject(   year, 
                                                semester, 
                                                subj, 
                                                specific_course_nums,
                                                columns_to_keep=columns_to_keep,
                                                add_semester_year_cols=True)
            # Add to list of dataframes
            all_data_frames_list.append(semester_dataframe)

    # Concatenate all dataframes
    all_data_pd = pd.concat(all_data_frames_list, axis=0)
    
    # Reset index
    all_data_pd.reset_index(drop=True, inplace=True)  

    # Return dataframe
    return all_data_pd