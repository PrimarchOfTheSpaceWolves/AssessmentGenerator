import MasterSchedule as ms
import Banner as bn
import Excel as ex
import pandas as pd
# Need to install package xlsxwriter 

COLUMNS_TO_KEEP = [
    "CRN",
    "Subj",
    "Crs",
    "Sec",
    "Title",
    "ENL",
    "Building",
    "Room",
    "Time",
    "Days",
    "Instructor"    
]

def __load_objectives(filename):
    # Load as Excel
    assess_goals = pd.read_excel(filename)
    # Return
    return assess_goals

def __combine_subj_crs(dataframe):
    dataframe["Crs"] = dataframe["Crs"].astype(str)
    dataframe.insert(3, "Subj_Crs", dataframe["Subj"].str.cat(dataframe["Crs"], sep =" "))
    dataframe = dataframe.drop(['Subj'], axis = 1)
    dataframe = dataframe.drop(['Crs'], axis = 1)
    return dataframe

def extract_unique_courses(objectives_pd):
    # Code snippet from: https://stackoverflow.com/questions/35268817/unique-combinations-of-values-in-selected-columns-in-pandas-data-frame-and-count
    chosen_cols = ['Subj','Crs']
    unique_courses = objectives_pd.groupby(chosen_cols).size().reset_index().rename(columns={0:'count'})
    unique_courses.sort_values(chosen_cols, axis=0, inplace=True, ignore_index=True)
    #print(unique_courses)
    
    subjects= {}
    
    for i in range(len(unique_courses)):
        subj = str(unique_courses["Subj"].iloc[i])
        crn = str(unique_courses["Crs"].iloc[i])
        
        if subj in subjects:
            subjects[subj].append(crn)
        else:
            subjects[subj] = [crn]
            
    #print(subjects)
            
    return subjects
    
def create_assessment_sheets(year_semester_list,                            
                            objectives_filename,
                            username,
                            password):
    
    # Load up objectives
    objectives_pd = __load_objectives(objectives_filename)
    
    # Get unique courses from objectives
    subjects = extract_unique_courses(objectives_pd)

    # Get courses
    courses_pd = ms.get_courses(    year_semester_list,
                                    subjects,                 
                                    columns_to_keep=COLUMNS_TO_KEEP)       

    # Create combined Subject + Course Number for key
    courses_pd = __combine_subj_crs(courses_pd)
    objectives_pd = __combine_subj_crs(objectives_pd)

    # Create grade dataframe
    grade_pd = courses_pd.copy()
    grade_pd = bn.get_grades(username=username, password=password, grade_pd=grade_pd)
    
    # Insert W/I/IP/EX values
    w_col_name = "W, I, IP, EX Count"
    courses_pd[w_col_name] = grade_pd[["W","I","IP","EX"]].sum(axis=1)       
    
    # Join two datasets
    all_data_pd = courses_pd.join(objectives_pd.set_index('Subj_Crs'), on='Subj_Crs')
    
    # Add formula for assessment outcome
    all_data_pd["Assessment Outcome"] = '=100*INDIRECT("Q" & ROW())/(INDIRECT("G" & ROW()) - INDIRECT("R" & ROW()))'
    
    # Mov W/I/IP/EX column
    w_col = all_data_pd.pop(w_col_name)
    all_data_pd.insert(17, w_col_name, w_col)
    
    # Return final results
    return all_data_pd, grade_pd

def generate_output_sheet_name(year_semester_list):
    sheet_name = ""
    for semester in year_semester_list:
        semester = semester.replace(" ", "-")
        if len(sheet_name) > 0:
            sheet_name += "_"
        sheet_name += semester
    return sheet_name

def generate_output_filename(year_semester_list):
    output_filename = generate_output_sheet_name(year_semester_list)
    output_filename += "_ASSESSMENT.xlsx"
    return output_filename

def save_assessment_sheets(year_semester_list, assess_pd, grade_pd):
    # Get filename
    output_filename = generate_output_filename(year_semester_list)

    # Create Excel writer
    writer = pd.ExcelWriter(output_filename)

    # Convert to Excel 
    assess_sheet_name = "Assessment" # generate_output_sheet_name(year_semester_list)
    grade_sheet_name = "Grades"
    assess_pd.to_excel(writer, sheet_name=assess_sheet_name, index=False)
    grade_pd.to_excel(writer, sheet_name=grade_sheet_name, index=False)

    # Auto-adjust column widths
    ex.auto_adjust_column_widths(writer, assess_pd, assess_sheet_name)
    ex.auto_adjust_column_widths(writer, grade_pd, grade_sheet_name)

    # Actually save Excel sheet
    writer.save()

def main():
    #year = ms.get_current_year()
    #sem = ms.get_current_semester()
    
    username="realemj"
    password=""
    
    year_semester_list = [ "2021 Fall", "2022 Spring"]
    #subjects = {
    #    "CS": ["108", "220", "240", "249", "330", "350", "370", "498"],
    #    "MAT": ["115", "413"]
    #}

    objectives_filename = "./OBJECTIVES/CS_OBJECTIVES.xlsx"

    # Generate assessment file
    all_data_pd, grade_pd = create_assessment_sheets(   year_semester_list,                                                        
                                                        objectives_filename,
                                                        username,
                                                        password)
    
    # Save assessment file
    save_assessment_sheets(year_semester_list, all_data_pd, grade_pd)    
        
if __name__ == "__main__":
    main()
