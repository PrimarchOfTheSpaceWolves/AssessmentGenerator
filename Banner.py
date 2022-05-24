import datetime
from enum import unique
import urllib.request as ul
from bs4 import BeautifulSoup as soup
import numpy as np
import pandas as pd
import requests

SEMESTERS = {
    "Fall" : ("09", "12"),
    "Spring" : ("01", "05"),
    "Summer" : ("06", "08")
}

GRADE_COLUMNS = ["A", "B", "C", "D", "F", "S", "U", "I", "IP", "L", "EX", "W" ]

def get_term_code(year, semester):
    sem_num = SEMESTERS[semester][0]
    term_in = str(year) + sem_num
    return term_in

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

# Returns requests.Session if valid; None otherwise.
def get_authenticated_banner_session(username, password):
    
    s = requests.Session() 
    
    url = 'https://banner.sunypoly.edu/pls/prod/twbkwbis.P_ValLogin'
    data =  {
        'sid': username,
        'PIN' : password
    }

    x = s.post(url)
    x = s.post(url, data = data)
    
    if "Authorization Failure" in x.text:
        return None
    else:    
        return s
    
def __parse_html_table_row(rowdata, header=False):
    pagesoup = soup(rowdata, "html.parser")
    if header:
        delim = 'th'
    else:
        delim = 'td'
    itemlocator = pagesoup.findAll(delim)
    
    data_elems = []
    for item in itemlocator:        
        data_elems.append(item.text)
        
    return data_elems    

def parse_banner_html_table_data(htmltext, tablename):
    try:         
        # Find the text "Summary Class List"
        second_half = str(htmltext).split(tablename)[1]
        second_half = second_half.split("</table>")[0]

        # Find each student
        pagesoup = soup(second_half, "html.parser")
        itemlocator = pagesoup.findAll('tr')

        column_names = []    
        all_data = []
        first_row = True    
                
        for item in itemlocator:
            stritem = str(item)
            #print("** " + stritem + "**")
                                    
            #if "<tr>" in stritem:
            # New row to process
            if first_row:
                column_names = __parse_html_table_row(stritem, header=True)                          
            else:
                data_cells = __parse_html_table_row(stritem, header=False)                                                       
                all_data.append(data_cells[:len(column_names)])
                                            
            first_row = False            
            
        # Convert all data to strings
        all_data_array = np.array(all_data).astype(str)
                
        # Convert to a Pandas dataframe
        all_data_pd = pd.DataFrame(all_data_array, columns=column_names).astype(str)
            
        # Clean up html &amp; text
        all_data_pd.replace("&amp;", "&", inplace=True, regex=True)

        # Remove NANs
        all_data_pd.fillna('',inplace=True)
    except:        
        all_data_pd = None
        
    return all_data_pd        
    
def get_banner_course_grades(session, year, semester, crn):
    
    term = get_term_code(year, semester)

    url="https://banner.sunypoly.edu/pls/prod/bwlkfcwl.P_FacClaListSum"    
    data = {
        "term" : term,
        "crn" : str(crn)
    }    
    x = session.post(url, data = data)
    
    all_data_pd = parse_banner_html_table_data(x.text, "Summary Class List")
            
    return all_data_pd    

def summarize_grade_counts(year, semester, crn, grades):
    if grades is not None:
        # Remove plus and minuses
        grades['Final'].replace("\+", "", inplace=True, regex=True)
        grades['Final'].replace("\-", "", inplace=True, regex=True)
        
        # Get unique counts
        summary = grades['Final'].value_counts()
        summary = summary.to_frame()
        
        # Join with unique grade possibilities
        unique_grades = pd.DataFrame(GRADE_COLUMNS, columns=["LETTER"])    
        unique_grades = unique_grades.set_index("LETTER")        
        summary = unique_grades.join(summary, how="left")
        
        # Remove NANs
        summary.fillna(0,inplace=True)
    else:
        # Make dummy table with empty values
        unique_grades = pd.DataFrame(GRADE_COLUMNS, columns=["LETTER"])    
        unique_grades = unique_grades.set_index("LETTER")    
        summary = unique_grades 
        summary['Final'] = 0               
    
    # Change to type int
    summary['Final'] = summary['Final'].astype(int)
    
    # Transpose to turn into row
    summary = summary.T
    
    # Add year, semester, and crn columns
    summary.insert(0, "Year", year)
    summary.insert(1, "Semester", semester)
    summary.insert(2, "CRN", crn)    
    
    # Reset index
    summary.reset_index(drop=True, inplace=True)                
    #print(summary)
    
    return summary
    
def get_grades(username, password, grade_pd):
    
    # Reset index
    grade_pd.reset_index(drop=True, inplace=True)  
    
    # Insert initial grade columns
    for grade in GRADE_COLUMNS:
        grade_pd.insert(len(grade_pd.columns), grade, 0)
        grade_pd[grade] = grade_pd[grade].astype(int)         
    
    # Get session
    s = get_authenticated_banner_session(username, password)
    
    # Get years, semesters, and crns
    all_years = grade_pd["Year"]
    all_semesters = grade_pd["Semester"]
    all_crns = grade_pd["CRN"]
    
    # For each row...
    for i in range(len(all_years)):
        # Get one row
        year = all_years.iloc[i]
        semester = all_semesters.iloc[i]
        crn = all_crns.iloc[i]
                
        # Get grades
        # TODO: Async calls?
        grades = get_banner_course_grades(s, year, semester, crn)
        grades = summarize_grade_counts(year, semester, crn, grades)
                
        # Insert values
        for col_name in GRADE_COLUMNS:
            grade_pd.at[i, col_name] = int(grades[col_name].iloc[0])
        
        #print(grade_pd.to_string())          
    
    grade_pd.insert(len(grade_pd.columns), "TOTAL", '=SUM(INDIRECT("M" & ROW() & ":X" & ROW()))')
    grade_pd.insert(len(grade_pd.columns), "ENL REPEATED", grade_pd["ENL"])
    
    # Reset index again
    grade_pd.reset_index(drop=True, inplace=True)  
    
    return grade_pd
        
def main():
    sess = get_authenticated_banner_session(username="realemj", password="dfdf")   
    if sess is None:
        print("HELP!")
        
if __name__ == "__main__":
    main()