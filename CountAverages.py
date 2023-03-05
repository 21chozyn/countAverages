import openpyxl

# Load the candidate information from the Excel file
wb = openpyxl.load_workbook('candidate_info.xlsx')
ws = wb.active
count = 0
ws3 = wb["Gen"]

#create variables for storing data
totalCandidatesCounted = 0
notCounted = 0
countRange0to49 = {"courseworkAverage":0,"examinationAverage":0}
countRange50to59 = {"courseworkAverage":0,"examinationAverage":0}
countRange60to69 = {"courseworkAverage":0,"examinationAverage":0}
countRange70to79 = {"courseworkAverage":0,"examinationAverage":0}
countRange80to100 = {"courseworkAverage":0,"examinationAverage":0}
candidatesWithoutValues = []
# Find the row in the Excel file that matches the last name
#1679 is number of rows i need to count
for rowNum in range(1,679):
    #to check if row has candidate name
    if ((ws3[f'A{rowNum}']).value == None) or (ws3[f'A{rowNum}']).value.startswith("2") == False:
        continue
    #to check if average column has values
    if ((ws3[f'G{rowNum}']).value == None) or (ws3[f'L{rowNum}']).value == None:
        notCounted+=1
        continue
    try:
        courseworkAv = sum([int((ws3[f'D{rowNum}']).value),int((ws3[f'E{rowNum}']).value),int((ws3[f'F{rowNum}']).value)])/3
        courseworkAv = round(courseworkAv)
        examinationsAv = sum([int((ws3[f'I{rowNum}']).value),int((ws3[f'J{rowNum}']).value),int((ws3[f'K{rowNum}']).value)])/3
        examinationsAv = round(examinationsAv)

    except:
        #to add candidate with missing data
        name=(ws3[f'B{rowNum}']).value
        candidatesWithoutValues.append(f'{rowNum}, {name}')
        notCounted+=1
        continue
    totalCandidatesCounted += 1
    print(totalCandidatesCounted)
    
    #adding the variables accordingly
    if courseworkAv in range(0,50):
        countRange0to49["courseworkAverage"] += 1
    elif courseworkAv in range(50,60):
        countRange50to59["courseworkAverage"] += 1
    elif courseworkAv in range(60,70):
        countRange60to69["courseworkAverage"] +=1
    elif courseworkAv in range(70,80):
        countRange70to79["courseworkAverage"] +=1
    elif courseworkAv in range(80,100):
        countRange80to100["courseworkAverage"] += 1
    
    if examinationsAv in range(0,50):
        countRange0to49["examinationAverage"] += 1
    elif examinationsAv in range(50,60):
        countRange50to59["examinationAverage"] += 1
    elif examinationsAv in range(60,70):
        countRange60to69["examinationAverage"] +=1
    elif examinationsAv in range(70,80):
        countRange70to79["examinationAverage"] +=1
    elif examinationsAv in range(80,100):
        countRange80to100["examinationAverage"] += 1

print("Number of students counted : ", totalCandidatesCounted)
print("In range 0 to 49 inclusive, " ,countRange0to49["courseworkAverage"]," for coursework average. ", countRange0to49["examinationAverage"]," for examination Average")
print("In range 50 to 59 inclusive, " ,countRange50to59["courseworkAverage"]," for coursework average. ", countRange50to59["examinationAverage"]," for examination Average")
print("In range 60 to 69 inclusive, " ,countRange60to69["courseworkAverage"]," for coursework average. ", countRange60to69["examinationAverage"]," for examination Average")
print("In range 70 to 79 inclusive, " ,countRange70to79["courseworkAverage"]," for coursework average. ", countRange70to79["examinationAverage"]," for examination Average")
print("In range 80 to 100 inclusive, " ,countRange80to100["courseworkAverage"]," for coursework average. ", countRange80to100["examinationAverage"]," for examination Average")
print(notCounted,"Candidates were not counted, no value in their cells")          

