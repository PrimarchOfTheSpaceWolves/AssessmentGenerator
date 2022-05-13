import MasterSchedule as ms
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

GRADE_COLUMNS = ["A", "B", "C", "D", "F", "S", "U", "I/IP/L" ]

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

def create_assessment_sheet(year_semester_list,
                            subjects,
                            objectives_filename):
    # Get courses
    courses_pd = ms.get_courses(  year_semester_list,
                                subjects,                 
                                columns_to_keep=COLUMNS_TO_KEEP)
    
    # Load up objectives
    objectives_pd = __load_objectives(objectives_filename)

    # Create combined Subject + Course Number for key
    courses_pd = __combine_subj_crs(courses_pd)
    objectives_pd = __combine_subj_crs(objectives_pd)

    # Join two datasets
    all_data_pd = courses_pd.join(objectives_pd.set_index('Subj_Crs'), on='Subj_Crs')
    
    # Add formula for assessment outcome
    all_data_pd["Assessment Outcome"] = '=100*INDIRECT("Q" & ROW())/INDIRECT("G" & ROW())'
    
    # Add grading columns
    for grade in GRADE_COLUMNS:
        all_data_pd.insert(len(all_data_pd.columns), "GRADE: " + grade, 0)

    # Return final result
    return all_data_pd

def generate_output_filename(year_semester_list):
    output_filename = ""
    for semester in year_semester_list:
        semester = semester.replace(" ", "-")
        if len(output_filename) > 0:
            output_filename += "_"
        output_filename += semester
    output_filename += "_ASSESSMENT.xlsx"
    return output_filename

def main():
    #year = ms.get_current_year()
    #sem = ms.get_current_semester()
    
    year_semester_list = [ "2021 Fall", "2022 Spring"]
    subjects = {
        "CS": ["108", "220", "240", "249", "330", "350", "370", "498"],
        "MAT": ["115", "413"]
    }

    objectives_filename = "OBJECTIVES.xlsx"

    # Generate assessment file
    all_data_pd = create_assessment_sheet(  year_semester_list,
                                            subjects,
                                            objectives_filename)
    
    # Save assessment file
    output_filename = generate_output_filename(year_semester_list)
    all_data_pd.to_excel(output_filename, index=False)
        
if __name__ == "__main__":
    main()
