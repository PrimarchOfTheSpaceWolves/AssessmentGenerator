import Assess
import Banner as bn
import os
import io
import pickle
import threading
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import scrolledtext

CACHE_FILENAME = 'cache.pkl'
errorLogGUI = None
doneProcessing = True
saveData = {}  
YEAR_OFFSET = 7

def setDefaultCached(key, defaultVal):
    global saveData
    if key not in saveData:
        saveData[key] = defaultVal

def loadCachedData():
    global saveData
    if os.path.exists(CACHE_FILENAME):      
        with open(CACHE_FILENAME, 'rb') as f:
            saveData = pickle.load(f)   

    # Set default keys if necessary
    setDefaultCached("objFile", "")
    setDefaultCached("outputFolder", "")
    setDefaultCached("username", "")
    setDefaultCached("startSemester", "")
    setDefaultCached("endSemester", "")
    
    # Resave, just in case defaults were added
    saveCachedData()

def saveCachedData():
    global saveData
    with open(CACHE_FILENAME, 'wb') as f:
        pickle.dump(saveData, f, 0) #pickle.HIGHEST_PROTOCOL)
        
def logToGUI(*args, **kwargs):
    output = io.StringIO()
    print(*args, file=output, **kwargs)
    contents = output.getvalue()
    errorLogGUI.insert(INSERT, contents)    
        
def runAssessment(year_semester_list, objectives_filename, username, password):

    try:    
        global doneProcessing
        doneProcessing = False
        
        logToGUI("Starting assessment...") 
        
        # Generate assessment file
        all_data_pd, grade_pd, summary_pd = Assess.create_assessment_sheets(    year_semester_list,                                                        
                                                                                objectives_filename,
                                                                                username,
                                                                                password)
        
        # Save assessment file
        output_filename = Assess.save_assessment_sheets(year_semester_list, all_data_pd, grade_pd, summary_pd)   
        
        logToGUI("Saved output file: " + output_filename)
        logToGUI("Assessment complete!")

    except Exception as e:
        logToGUI(e)
        logToGUI("Assessment FAILED!")

    doneProcessing = True
    
def performAssessment(year_semester_list, objectives_filename, username, password):
    if doneProcessing:
        try:
            extractThread = threading.Thread(target=runAssessment, args=(year_semester_list, objectives_filename, username, password,))
            extractThread.start()             
        except:
            logToGUI("ERROR: Assessment cannot be performed!")   
    else:
        logToGUI("ERROR: Please wait for processing to complete!")  

def initGUI():
    # Create main window
    window = Tk()
    window.title("Assessment Generator")
    
    # Try to load previously cached data    
    loadCachedData()
        
    # Top and bottom frame
    topFrame = Frame(window)
    topFrame.pack(fill=BOTH)
    bottomFrame = Frame(window)
    bottomFrame.pack(fill=BOTH)
    
    # ********************************************
    # OBJECTIVES FILE
    # ********************************************

    # Add label and button for opening objectives file
    objFileFrame = Frame(topFrame)
    objFileFrame.pack(fill=X)
    
    # Set up text variable 
    objFileText = StringVar()
    objFileText.set(saveData["objFile"])     
    
    # Set up label
    objFileLabel = Label(objFileFrame, textvariable=objFileText) 
    objFileLabel.pack(side=LEFT, padx=5, pady=5)   
    
    # Define function for opening objectives file
    def openObjectivesFile():
        # Get current text
        curFile = objFileText.get()
        
        # Get last directory
        curDir = os.path.dirname(curFile)
                
        # Set allowable filetypes
        filetypes = (
            ('Excel files', '*.xlsx'),
            ('All files', '*.*')
        )

        # Open file dialog
        objFile = filedialog.askopenfilename(title="Open Objectives File",
                                            initialdir=curDir,
                                            filetypes=filetypes)
        
        if len(objFile) > 0:
            # Update text
            objFileText.set(objFile)

            # Save cached data
            global saveData        
            saveData["objFile"] = objFile
            saveCachedData()

    objFileLoadButton = Button(objFileFrame, text="Open Objectives File", command=openObjectivesFile)
    objFileLoadButton.pack(side=RIGHT, padx=5, pady=5) 
    
    # ********************************************
    # USERNAME AND PASSWORD
    # ********************************************
    usernameFrame = Frame(topFrame)
    usernameFrame.pack(fill=X)    
    usernameText = StringVar()    
    usernameText.set(saveData["username"])   
      
    usernameLabel = Label(usernameFrame, text="SITNET ID:") 
    usernameLabel.pack(side=LEFT, padx=5, pady=5)   
    
    def save_username(value):
        # Save cached data
        global saveData        
        saveData["username"] = usernameText.get()
        saveCachedData()
    
    usernameInput = Entry(usernameFrame, textvariable=usernameText)
    usernameInput.bind('<Return>',save_username)
    usernameInput.pack(side=RIGHT, padx=5, pady=5)   
    
    passwordFrame = Frame(topFrame)
    passwordFrame.pack(fill=X)    
    passwordText = StringVar()    
          
    passwordLabel = Label(passwordFrame, text="Password:") 
    passwordLabel.pack(side=LEFT, padx=5, pady=5)   
    
    passwordInput = Entry(passwordFrame, textvariable=passwordText, show="*")
    passwordInput.pack(side=RIGHT, padx=5, pady=5)  
    
    # ********************************************
    # SEMESTER BOUNDS
    # ********************************************
    
    semesterFrame = Frame(topFrame)
    semesterFrame.pack(fill=X) 
        
    start_semester_text = StringVar()
    end_semester_text = StringVar()         
    
    def update_semesters():
        # Save cached data
        global saveData        
        saveData["username"] = usernameText.get()
        saveCachedData()
        
    def populate_semester_combo(combo):
        # Get current year and semester
        year = int(bn.get_current_year())
        semester = bn.get_current_semester()
                
        # Get semester keys
        sem_keys = list(bn.SEMESTERS.keys())
        
        # Find current semester
        sem_index = sem_keys.index(semester)
        num_semesters = len(sem_keys)
        
        # Generate list going backwards
        all_semesters = []
        for i in range(YEAR_OFFSET*num_semesters):
            # Get current semester
            semester = sem_keys[sem_index]
            # Get current semester string
            semester_string = str(year) + " " + semester
            # Add to list
            all_semesters.append(semester_string)
            # Decrement semester index 
            sem_index -= 1
            if sem_index < 0:
                sem_index = num_semesters - 1
            # Did we just leave Spring?
            if semester == "Spring":
                # Decrement year
                year -= 1  
        
        combo['values'] = all_semesters
        
        # Make sure there are no custom values
        combo['state'] = 'readonly'
            
    Label(semesterFrame, text="Start Semester:").pack(side=LEFT, padx=5, pady=5) 
    start_semester_combo = ttk.Combobox(semesterFrame, textvariable=start_semester_text) 
    populate_semester_combo(start_semester_combo)
    def save_start_semester(box):
        # Save cached data
        global saveData        
        saveData["startSemester"] = start_semester_text.get()        
        saveCachedData()
        
    start_semester_combo.bind('<<ComboboxSelected>>', save_start_semester)
    start_semester_combo.set(saveData["startSemester"])
    start_semester_combo.pack(side=LEFT, padx=5, pady=5) 
      
    Label(semesterFrame, text="End Semester:").pack(side=LEFT, padx=5, pady=5)     
    end_semester_combo = ttk.Combobox(semesterFrame, textvariable=end_semester_text)  
    populate_semester_combo(end_semester_combo)
    def save_end_semester(box):
        # Save cached data
        global saveData        
        saveData["endSemester"] = end_semester_text.get()        
        saveCachedData()
    end_semester_combo.bind('<<ComboboxSelected>>', save_end_semester)  
    end_semester_combo.set(saveData["endSemester"])  
    end_semester_combo.pack(side=LEFT, padx=5, pady=5) 
        
    # ********************************************
    # OUTPUT FOLDER
    # ********************************************
    
    outputFolderFrame = Frame(topFrame)
    outputFolderFrame.pack(fill=X)
    
    outputFolderText = StringVar()
    outputFolderText.set(saveData["outputFolder"])     
    
    # Set up label
    outputFolderLabel = Label(outputFolderFrame, textvariable=outputFolderText) 
    outputFolderLabel.pack(side=LEFT, padx=5, pady=5)   
    
    # Define function for opening output folder
    def openOutputFolder():
        # Get current text
        curDir = outputFolderText.get()
        
        # Open folder dialog
        outputFolder = filedialog.askdirectory(title="Open Output Folder",
                                            initialdir=curDir)
        
        if len(outputFolder) > 0:
            # Update text
            outputFolderText.set(outputFolder)

            # Save cached data
            global saveData        
            saveData["outputFolder"] = outputFolder
            saveData["username"] = usernameText.get()
            saveCachedData()
            
            # Get semester list
            semester_values = list(start_semester_combo["values"])
            start_index = semester_values.index(start_semester_text.get())
            end_index = semester_values.index(end_semester_text.get())  
            year_semester_list = semester_values[end_index:(start_index+1)]
            year_semester_list.reverse()
            
            # Check values
            if len(year_semester_list) > 0:            
                # Do assessment!            
                performAssessment(year_semester_list=year_semester_list, 
                                objectives_filename=objFileText.get(), 
                                username=usernameText.get(), 
                                password=passwordText.get())
            else:
                logToGUI("ERROR: Invalid semester range!")
                logToGUI("Assessment FAILED!")

    outputFolderButton = Button(outputFolderFrame, text="Save Output Assessment File", command=openOutputFolder)
    outputFolderButton.pack(side=RIGHT, padx=5, pady=5) 
    
    # ********************************************
    # ERROR TEXT AREA
    # ********************************************
        
    global errorLogGUI
    errorLogGUI = scrolledtext.ScrolledText(bottomFrame) #width=40,height=10)
    errorLogGUI.pack(side=TOP, fill=BOTH, padx=5, pady=5)  
        
    # Return window object    
    return window

def main():
    window = initGUI()
    window.mainloop()

if __name__ == "__main__":
    main()
