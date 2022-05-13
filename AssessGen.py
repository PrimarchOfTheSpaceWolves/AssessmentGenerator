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

def save_assessment_sheet(year_semester_list, dataframe):
    # Get filename
    output_filename = generate_output_filename(year_semester_list)

    # Create Excel writer
    writer = pd.ExcelWriter(output_filename)

    # Convert to Excel 
    sheet_name = generate_output_sheet_name(year_semester_list)
    dataframe.to_excel(writer, sheet_name=sheet_name, index=False)

    # Code from: https://stackoverflow.com/questions/45985358/how-to-wrap-text-for-an-entire-column-using-pandas
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
    wrap_format = workbook.add_format({'text_wrap': True})

    # Code from: https://towardsdatascience.com/how-to-auto-adjust-the-width-of-excel-columns-with-pandas-excelwriter-60cee36e175e
    # Auto-adjust columns' width
    # NOTE: Need XlsxWriter package!
    MAX_COLUMN_WIDTH = 30
    OFFSET = 2
    for column in dataframe:
        data_len = dataframe[column].astype(str).map(len).max() + OFFSET
        name_len = len(column) + OFFSET
        
        # Column width should NOT exceed max
        column_width = min(data_len, MAX_COLUMN_WIDTH)
        # ...unless the name is longer
        column_width = max(column_width, name_len)
        
        # Get index of column
        col_idx = dataframe.columns.get_loc(column)    

        # Set width (while also making sure wrapping is enabled)     
        writer.sheets[sheet_name].set_column(col_idx, col_idx, column_width, wrap_format)

    # Actually save Excel sheet
    writer.save()

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
    save_assessment_sheet(year_semester_list, all_data_pd)    
        
if __name__ == "__main__":
    main()
