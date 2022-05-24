import datetime
import urllib.request as ul
from bs4 import BeautifulSoup as soup
import numpy as np
import pandas as pd

SEMESTERS = {
    "Fall" : ("09", "12"),
    "Spring" : ("01", "05"),
    "Summer" : ("06", "08")
}

GRADE_COLUMNS = ["A", "B", "C", "D", "F", "S", "U", "I", "IP", "L", "EX", "W" ]

# TODO: Revamp using requests library and soup subparsing
def __split_columns(data, begin, to_remove):
    for t in to_remove:
        data = data.replace(t, "")            
    tokens = data.split(begin)      
    tokens = tokens[1:]      
    return tokens 

def get_term_code(year, semester):
    sem_num = SEMESTERS[semester][0]
    term_in = str(year) + sem_num
    return term_in

def get_subject(year, semester, subj, 
                specific_course_nums=None,
                columns_to_keep=None,
                add_semester_year_cols=False):
        
    term_in = get_term_code(year, semester)

    base_url="https://banner.sunypoly.edu/pls/prod/swssschd.P_ShowSchd" 
    base_url += "?"
    base_url += "term_in=" + term_in
    base_url += "&"
    base_url += "disc_in=" + subj
    
    req = ul.Request(base_url, headers={'User-Agent': 'Mozilla/5.0'})
    client = ul.urlopen(req)
    htmldata = client.read()
    client.close()
    
    # Find the text "Master Schedule"
    second_half = str(htmldata).split("Master Schedule")[1]
    second_half = second_half.split("</table>")[0]

    # Find each course
    pagesoup = soup(second_half, "html.parser")
    itemlocator = pagesoup.findAll('tr')

    column_names = []    
    all_data = []
    first_row = True    
            
    for item in itemlocator:
        stritem = str(item)
        #print("** " + stritem + "**")
                                   
        if "<tr>" in stritem:
            # New row to process
            if first_row:                
                column_names = __split_columns(stritem, "<th>", ["<tr>", "</th>", "</tr>", "\\n"])            
            else:
                data_cells = __split_columns(stritem, "<td>", ["<tr>", "</td>", "</tr>", "\\n"])    
                
                if specific_course_nums is not None:
                    num_index = column_names.index("Crs")
                    crs_num = data_cells[num_index]
                    if crs_num in specific_course_nums:
                        all_data.append(data_cells)
                else:
                    all_data.append(data_cells)
                                            
            first_row = False            
           
    all_data_array = np.array(all_data)

    # Convert to a Pandas dataframe
    all_data_pd = pd.DataFrame(all_data_array, columns=column_names).astype(str)
    
    # If asked, only keep certain columns
    if columns_to_keep is not None:
        all_data_pd = all_data_pd[columns_to_keep]

    # Clean up html &amp; text
    all_data_pd.replace("&amp;", "&", inplace=True, regex=True)

    # Remove NANs
    all_data_pd.fillna('',inplace=True)

    # Add semester and year columns?
    if add_semester_year_cols:
        all_data_pd.insert(0, "Year", year)
        all_data_pd.insert(1, "Semester", semester)

    return all_data_pd        

def get_current_year():
    current_year = int(datetime.datetime.now().date().strftime("%Y"))
    return current_year      

def get_current_semester():
    current_month = int(datetime.datetime.now().date().strftime("%m"))
    for sem in SEMESTERS:
        month_range = SEMESTERS[sem]
        start_month = int(month_range[0])
        end_month = int(month_range[1])
        if start_month <= current_month <= end_month:
            return sem

    raise ValueError("Unable to find valid semester!") 
   
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

    # Return dataframe
    return all_data_pd