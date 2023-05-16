import MasterSchedule as ms
import Banner as bn
import Excel as ex
import pandas as pd

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

W_COL_NAME = "W, I, IP, EX Count"

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

def create_summary_dataframe(year_semester_list, all_data_pd, objectives_pd):
    
    # Create column name
    year_column_name = ""
    for val in year_semester_list:
        if len(year_column_name) == 0:
            year_column_name = val
        else:
            year_column_name += " - " + val
        
    # Get unique courses
    unique_courses = objectives_pd.groupby("Subj_Crs").size().reset_index()["Subj_Crs"].astype(str)
        
    # Create summary dataframe
    summary_data = objectives_pd.copy()    
    summary_data = summary_data[["Subj_Crs", 
                                "Outcome ID", 
                                "Assessment Outcome"]]
    summary_data = summary_data.rename(columns={"Assessment Outcome":year_column_name}) 
    
    # Get Excel column names for relevant columns
    ex_count = ex.convert_pandas_col_to_excel_col(all_data_pd, "Count Meeting Outcome")
    ex_enrolled = ex.convert_pandas_col_to_excel_col(all_data_pd, "ENL")
    ex_ws = ex.convert_pandas_col_to_excel_col(all_data_pd, W_COL_NAME)
    
    ex_sheet = "Assessment"
    ex_count    = ex_sheet + "!" + ex_count
    ex_enrolled = ex_sheet + "!" + ex_enrolled
    ex_ws       = ex_sheet + "!" + ex_ws
    
    
    # Define local formula-generation function
    def get_formula_for_sum(column, indices):
        part = "(0"
        for i in indices:
            # NOTE: +2 is because data doesn't start in Excel until second row (which is 1-based indicing)
            part += "+" + column + str(i+2)
        part += ")"
        return part
    
    # Create columns for enrolled, count meeting, W students
    total_enrolled_col = year_column_name + " (Total Enrolled)"
    w_students_col = year_column_name + " (W/I/IP/EX Students)"
    count_meeting_col = year_column_name + " (Count Meeting Outcome)"
    
    summary_data[total_enrolled_col] = 0
    summary_data[w_students_col] = 0    
    summary_data[count_meeting_col] = 0
    
    # Move assessment outcome column
    a_col = summary_data.pop(year_column_name)
    year_column_name += " (Assessment Outcome)"
    summary_data.insert(len(summary_data.columns), year_column_name, a_col)
    
    # Get Excel reference for THIS sheet
    ex_local_enrolled = ex.convert_pandas_col_to_excel_col(summary_data, total_enrolled_col)
    ex_local_ws = ex.convert_pandas_col_to_excel_col(summary_data, w_students_col)
    ex_local_count = ex.convert_pandas_col_to_excel_col(summary_data, count_meeting_col)
                
    # For each course...
    for course in unique_courses:        
        # Get unique objectives
        objective_data = objectives_pd[objectives_pd["Subj_Crs"] == course]
        
        # For each objective
        for obj in objective_data["Outcome ID"]:
            # Get specific course and objective data
            #print(all_data_pd.columns)
            specific_obj_data = all_data_pd.loc[(all_data_pd["Subj_Crs"] == course) 
                                                & (all_data_pd["Outcome ID"] == obj)]  
            
            #print("INDICES:", course, obj, specific_obj_data.index) 
            #print(specific_obj_data)         
                        
            # Get formula for counts
            ex_count_part = get_formula_for_sum(ex_count, specific_obj_data.index)
            ex_enrolled_part = get_formula_for_sum(ex_enrolled, specific_obj_data.index)
            ex_ws_part = get_formula_for_sum(ex_ws, specific_obj_data.index)
                        
            # Get final formula            
            # Need "Count Meeting Outcome" / ("ENL" - W_COL_NAME)
            #ex_final_formula = "=100*" + ex_count_part + "/" + "(" + ex_enrolled_part + "-" + ex_ws_part + ")" 
            
            indirect_enrolled = 'INDIRECT("' + ex_local_enrolled + '" & ROW())'
            indirect_count = 'INDIRECT("' + ex_local_count + '" & ROW())'
            indirect_ws = 'INDIRECT("' + ex_local_ws + '" & ROW())'
            
            ex_final_formula = "=100*" + indirect_count + "/" + "(" + indirect_enrolled + "-" + indirect_ws + ")" 
                                   
            # Write to table 
            summary_data.loc[
                (summary_data["Subj_Crs"] == course) & 
                (summary_data["Outcome ID"] == obj),
                total_enrolled_col] = "=" + ex_enrolled_part     
            
            summary_data.loc[
                (summary_data["Subj_Crs"] == course) & 
                (summary_data["Outcome ID"] == obj),
                w_students_col] = "=" + ex_ws_part     
            
            summary_data.loc[
                (summary_data["Subj_Crs"] == course) & 
                (summary_data["Outcome ID"] == obj),
                count_meeting_col] = "=" + ex_count_part       
            
            summary_data.loc[
                (summary_data["Subj_Crs"] == course) & 
                (summary_data["Outcome ID"] == obj),
                year_column_name] = ex_final_formula               
    
    summary_data.reset_index(drop=True, inplace=True)  
    #print(summary_data)
    
    return summary_data
    
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
    
    # Grab copy of objectives
    metrics_pd = objectives_pd.copy() 
    
    # Prepare metrics dataframe    
    metrics_pd = __combine_subj_crs(metrics_pd)
    metrics_pd.drop(columns=['Indirect Measure Used', 'Count Meeting Outcome', 
                             'Assessment Outcome', 'Action Taken', 'Additional Comments and Notes'], 
                    inplace=True)
    metrics_pd=metrics_pd[metrics_pd.columns[[1,0,2,3]]] # Changing column order

    # Create combined Subject + Course Number for key
    courses_pd = __combine_subj_crs(courses_pd)
    objectives_pd = __combine_subj_crs(objectives_pd)

    # Create grade dataframe
    grade_pd = courses_pd.copy()
    grade_pd = bn.get_grades(username=username, password=password, grade_pd=grade_pd)
    
    # Insert W/I/IP/EX values    
    courses_pd[W_COL_NAME] = grade_pd[["W","I","IP","EX"]].sum(axis=1)       
    
    # Join two datasets
    all_data_pd = courses_pd.join(objectives_pd.set_index('Subj_Crs'), on='Subj_Crs')
    
    # Add formula for assessment outcome
    all_data_pd["Assessment Outcome"] = '=100*INDIRECT("Q" & ROW())/(INDIRECT("G" & ROW()) - INDIRECT("R" & ROW()))'
    
    # Mov W/I/IP/EX column
    w_col = all_data_pd.pop(W_COL_NAME)
    all_data_pd.insert(17, W_COL_NAME, w_col)
    
    # Reset all indices
    all_data_pd.reset_index(drop=True, inplace=True)       
    grade_pd.reset_index(drop=True, inplace=True)       
    
    # Create summary dataframe
    summary_pd = create_summary_dataframe(year_semester_list, all_data_pd, objectives_pd)
    
    # Clean out the Student Learning and Direct Measure columns, since they are on a separate sheet now
    all_cols = all_data_pd.columns.tolist()
    all_cols_data = all_cols[:]
    all_cols_data.remove('Student learning Outcome/Performance Indicator')
    all_cols_data.remove('Direct Measure Used')    
    all_data_pd = pd.DataFrame(all_data_pd[all_cols_data], columns=all_cols)    
    all_data_pd.fillna('',inplace=True)
            
    # Return final results
    return all_data_pd, metrics_pd, grade_pd, summary_pd

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

def save_assessment_sheets(year_semester_list, assess_pd, metrics_pd, grade_pd, summary_pd):
    # Get filename
    output_filename = generate_output_filename(year_semester_list)

    # Create Excel writer
    writer = pd.ExcelWriter(output_filename)

    # Convert to Excel 
    assess_sheet_name = "Assessment" # generate_output_sheet_name(year_semester_list)
    metrics_sheet_name = "Metrics"
    grade_sheet_name = "Grades"
    summary_sheet_name = "Summary"
    assess_pd.to_excel(writer, sheet_name=assess_sheet_name, index=False)
    metrics_pd.to_excel(writer, sheet_name=metrics_sheet_name, index=False)
    grade_pd.to_excel(writer, sheet_name=grade_sheet_name, index=False)
    summary_pd.to_excel(writer, sheet_name=summary_sheet_name, index=False)

    # Auto-adjust column widths
    ex.auto_adjust_column_widths(writer, assess_pd, assess_sheet_name)
    ex.auto_adjust_column_widths(writer, metrics_pd, metrics_sheet_name)
    ex.auto_adjust_column_widths(writer, grade_pd, grade_sheet_name)
    ex.auto_adjust_column_widths(writer, summary_pd, summary_sheet_name)

    # Actually save Excel sheet
    writer.save()
    
    return output_filename

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
    all_data_pd, metrics_pd, grade_pd, summary_pd = create_assessment_sheets(   year_semester_list,                                                        
                                                                    objectives_filename,
                                                                    username,
                                                                    password)
    
    # Save assessment file
    save_assessment_sheets(year_semester_list, all_data_pd, metrics_pd, grade_pd, summary_pd)    
        
if __name__ == "__main__":
    main()
