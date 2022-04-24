"""
Data Visualization for CTL.
Main features:
    1. Ability to run the script regardless of computer/directories; you can also simply
       just add new .xlsx files to the CTLData directory without having to edit any code here
       under the condition that it is CTLData->folder1->folder2->file.
    2. Flexibility with which file to parse (.xlsx now, but can be changed to any file on line 37)
    3. schoolData dict that can be modified to specifications

Installation/Running requirements:
Use python 3, you may need to download the following libraries:
    pip install openpyxl
    pip install matplotlib
    pip install numpy

To run: type "python3 data.py" in the CTLData directory.
"""
import matplotlib.pyplot as plt     # plotting
import pandas as pd                 # dataframes and .xlsx -> python objects
import os                           # for iterating over directories; general os functionality


wantedCol = "Stanford School"       # the column name that we will parse data from
schoolData = {                      # data dict storing frequency of schools
    "Business": 0,
    "Earth, Energy, and Environmental Sciences": 0,
    "Engineering": 0,
    "Education": 0,
    "Humanities and Sciences": 0,
    "Medicine": 0,
    "Law": 0,
    "Undergrad/NA": 0
}


# Gets the current working directory and returns a list of all the .xlsx files inside of the directory
def parseDir():
    directory = os.getcwd()
    filePaths = []
    for dir in os.listdir(directory):
        f1 = os.path.join(directory, dir)
        if os.path.isdir(f1):   # confirm this is a directory path
            for nextDir in os.listdir(f1):
                f2 = os.path.join(f1, nextDir)
                if os.path.isdir(f2):   # confirm this is a directory path
                    for fileName in os.listdir(f2):
                        if ".xlsx" in fileName:   # parse only .xlsx files
                            filePaths.append(os.path.join(f2, fileName))   # add the final file path to aa list
    return filePaths


# Function that takes in a list of strings and replaces all spaces with newline chars,
# used to wrap the text of tickmarks when plotting them using matplotlib
def wrap(names):
    for i in range(len(names)):
        newString = ""
        for x in range(len(names[i])):
            if names[i][x] == " ":
                newString += "\n"
            else:
                newString += names[i][x]
        names[i] = newString
    return names


def main():
    filePaths = parseDir()
    for file in filePaths:
        print("Reading file: " + file)
        df = pd.read_excel(file)

        # Only select the files w/ the "Stanford School" col, aka only
        # the files with info about the attendee's school
        if wantedCol in df:
            df = df[[wantedCol]]     # df now only contains the column wantedCol
            for row in df[wantedCol]:
                if (pd.isna(row)):   # if the entry is NAN, it's an undergrad/NA
                    schoolData["Undergrad/NA"] += 1
                else:
                    schoolData[row] += 1

    names = list(schoolData.keys())
    values = list(schoolData.values())
    names = wrap(names)                        # wrap names for the tickmarks
    plt.figure(figsize=(12, 6))                # resize the graph manually

    plt.bar(names, values, tick_label=names)   # plot the data as a bar graph, add labels
    plt.xlabel("School")
    plt.ylabel("No. of students attending")
    plt.title("School distribution of event attendance ")
    plt.tight_layout()                         # constrain the graph to the window
    print("Complete!")
    plt.show()

if __name__ == "__main__":
    main()
