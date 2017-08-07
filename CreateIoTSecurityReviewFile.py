import os.path
import xlwt

Phases = ['1. Evaluation des Risques', '2. HLRA - Risk Assessments', '3. Audits de Securite', '4. Tests Manuels', '5. Legal and Contracts', '6. Delivery and Quality']

ProjectsTables = [['1. Evaluation des Risques'], ['2. HLRA - Risk Assessments'], ['3. Audits de Securite'], ['4. Tests Manuels'], ['5. Legal and Contracts'], ['6. Delivery and Quality']]


NotStartedPhase = 'not started'
InProgressPhase = 'in progress'
DonePhase = 'done'
Separator = ' - '
exten = '.xls'
currentDir = '.'

NotStartedProjectsCounter = 0
InProgressProjectsCounter = 0
DoneProjectsCounter = 0

def getSubstringFromStringSeparatorAndIndex(string, separator, index):
    list = string.split(separator)
    objectName = ''
    try:
        objectName = list[index]
    except IndexError:
        objectName = getSubstringFromStringSeparatorAndIndex(string, separator, index-1)
        pass
    
    return objectName

def getPhaseNameFrom(dirpath):
    list = dirpath.split('\\')
    try:
        phaseName = list[1]
    except IndexError:
        pass

    return phaseName

def getProjectNameFrom(dirName):
    projectName = ''
    try:
        projectName = getSubstringFromStringSeparatorAndIndex(dirName, Separator, 1)
    except IndexError:
        pass

    return projectName

def getObjectNameFrom(dirName):
    list = dirname.split(Separator)
    objectName = ''
    try:
        objectName = getSubstringFromStringSeparatorAndIndex(dirName, Separator, 2)
    except IndexError:
        pass
    
    return objectName

def writeExcelFileFrom(projects):
    row=0
    column=0
    book = xlwt.Workbook(encoding="utf-8")

    sheet1 = book.add_sheet("Sheet 1")
    for projectsInPhase in projects:

        for projects in projectsInPhase:
            sheet1.write(column, row, projects)
            row=row+1
        column=column+1
        row=0
    book.save("trial.xls")


def writeCsvFileFrom(projects):
    fileObject = open('SuiviProjetIoT.csv', "w")

    for phases in ProjectsTables:
        for phase in phases:
            fileObject.write(phase + ',' + os.linesep)

    fileObject.write(os.linesep)

#Main function
for dirpath, dirnames, files in os.walk(currentDir):

    NotStartedObjectList = []
    InProgressObjectList = []
    DoneProjectList = []

    
    for dirname in dirnames:
        if dirname.lower().startswith(NotStartedPhase):
            currentPhase = getPhaseNameFrom(dirpath)
            phaseIndexInTables = Phases.index(currentPhase)
            ProjectsTables[phaseIndexInTables].append(getObjectNameFrom(dirname))

            #print('Phase : ' + currentPhase)
            #print('Project Name : ' + getProjectNameFrom(dirname))
            #print('Object Name : ' + getObjectNameFrom(dirname))
            NotStartedProjectsCounter = NotStartedProjectsCounter + 1

    for dirname in dirnames:
        if dirname.lower().startswith(InProgressPhase):
            currentPhase = getPhaseNameFrom(dirpath)
            phaseIndexInTables = Phases.index(currentPhase)
            ProjectsTables[phaseIndexInTables].append(getObjectNameFrom(dirname))

            #print('Phase : ' + getPhaseNameFrom(dirpath))
            #print('Project Name : ' + getProjectNameFrom(dirname))
            #print('Object Name : ' + getObjectNameFrom(dirname))
            InProgressProjectsCounter = InProgressProjectsCounter + 1

    
    for dirname in dirnames:
        if dirname.lower().startswith(DonePhase):
            currentPhase = getPhaseNameFrom(dirpath)
            phaseIndexInTables = Phases.index(currentPhase)
            ProjectsTables[phaseIndexInTables].append(getObjectNameFrom(dirname))

            #print('Phase : ' + getPhaseNameFrom(dirpath))
            #print('Project Name : ' + getProjectNameFrom(dirname))
            #print('Object Name : ' + getObjectNameFrom(dirname))
            DoneProjectsCounter = DoneProjectsCounter + 1

print('Review')
print('Not Started Projects : ' + str(NotStartedProjectsCounter))
print('In Progress Projects : ' + str(InProgressProjectsCounter))
print('Done Projets : ' + str(DoneProjectsCounter))

#writeCsvFileFrom(ProjectsTables)
writeExcelFileFrom(ProjectsTables)


print(ProjectsTables)