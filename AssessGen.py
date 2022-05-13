import MasterSchedule as ms

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
     
def main():
    #year = ms.get_current_year()
    #sem = ms.get_current_semester()
    
    year_semester_list = [ "2021 Fall", "2022 Spring"]
    subjects = {
        "CS": ["108", "220", "240", "249", "330", "350", "370", "498"],
        "MAT": ["115", "413"]
    }

    all_data_pd = ms.get_courses(  year_semester_list,
                                subjects,                 
                                columns_to_keep=COLUMNS_TO_KEEP)
    
    print(all_data_pd)
    
    all_data_pd.to_excel("ALL_COURSES.xlsx", index=False)
        
if __name__ == "__main__":
    main()
