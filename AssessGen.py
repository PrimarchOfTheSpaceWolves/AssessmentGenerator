import urllib.request as ul
from bs4 import BeautifulSoup as soup
import re
import datetime
import pandas as pd
import numpy as np

SEMESTERS = {
    "Fall" : "09",
    "Summer" : "06", 
    "Spring" : "01"
}

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

def split_columns(data, begin, to_remove):
    for t in to_remove:
        data = data.replace(t, "")            
    tokens = data.split(begin)      
    tokens = tokens[1:]      
    return tokens       

def get_courses(year, semester, subj, specific=None):
    
    sem_num = SEMESTERS[semester]
    term_in = str(year) + sem_num

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
                column_names = split_columns(stritem, "<th>", ["<tr>", "</th>", "</tr>", "\\n"])            
            else:
                data_cells = split_columns(stritem, "<td>", ["<tr>", "</td>", "</tr>", "\\n"])    
                
                if specific is not None:
                    num_index = column_names.index("Crs")
                    crs_num = data_cells[num_index]
                    if crs_num in specific:
                        all_data.append(data_cells)
                else:
                    all_data.append(data_cells)
                #print(stritem)
                #print(len(data_cells))
                            
            first_row = False            
           
    all_data_array = np.array(all_data)
    #print(all_data_array) 
    
    return all_data_array, column_names  
        
     
def main():
    current_year = int(datetime.datetime.now().date().strftime("%Y"))
    year = 2021 # current_year
    sem = "Fall"
    cs_data, column_names = get_courses(year, sem, subj="CS", specific=["108", "220", "240", "249", "330", "350", "370", "498"])    
    mat_data, _ = get_courses(year, sem, subj="MAT", specific=["115", "413"])    
    all_data_array = np.concatenate([cs_data, mat_data])        
    all_data_pd = pd.DataFrame(all_data_array, columns=column_names).astype(str)
    all_data_pd = all_data_pd[COLUMNS_TO_KEEP]
    all_data_pd.replace("&amp;", "&", inplace=True, regex=True)
    print(all_data_pd)
    
    all_data_pd.to_excel(str(year) + "_" + sem + "_COURSES.xlsx", index=False)
        
if __name__ == "__main__":
    main()
