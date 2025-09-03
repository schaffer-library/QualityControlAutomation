"""
Authored by James Gaskell

11/06/2024

Edited by: Mukundan Thanigaivelan

"""
import shutil
from typing import List
import pandas as pd
import re, csv, os, easygui

DEFAULT_COLORS = os.path.join(os.path.dirname(__file__), 'assets', 'defaultColors.csv')
COLOR_PALETTE = os.path.join(os.path.dirname(__file__), 'assets', 'errorColors.csv')
CACHED_COLORS = os.path.join(os.path.dirname(__file__), 'assets' , 'prevColors.txt')

class ScanFile():

    def __init__(self, filename, loc, date, pages):
        self.fileName = filename
        self.location = loc
        self.date = date
        self.filePath = None
        self.exists = False
        self.extent = pages # This is the extent read from the spreadsheet - not the actual extent

        # The order of these is important for the error returns
        # The file is failed or highlighted based on the last error that matches
        # The dictionary must go in descending order of importance
        # Order:
        # Date
        # Filename
        # DupFilename
        # Extent
        # Filesize
        # Existence
        self.errors = {'Date': False,
                       'Filename': False,
                       'DupFilename': False}
        
        self.failures = {
                        'Extent': False,
                        'Filesize': False,
                        'Existence': False}
    




class ExcelSheet():

    def __init__(self, sheetname):
        self.sheetName = sheetname
        self.errors = 0
        self.failures = 0
        self.fileList: List[ScanFile] = list()

    def getErrorRate(self):
        return (self.errors / len(self.fileList) * 100)
    
    def getFailureRate(self):
        return (self.failures / len(self.fileList) * 100)

    """Creates a list of ScanFile objects from the dataframe"""
    def createScanFileList(self, df):
        for _, row in df.iterrows():
            try:
                spreadsheet_extent = int(re.findall(r'\d+', row['extent (total page count including covers)'])[0])
            except:
                spreadsheet_extent = None
            record = ScanFile(row['Filename'], # --> Problem; program is looking for 
                                # column names Filename, physical_location, and date_created;
                                # if it doesn't find them, then it simply returns without
                                # doing anything
                                row['physical_location'],
                                row['date_created'] if not pd.isna(row['date_created']) else None,
                                spreadsheet_extent)
            
            self.fileList.append(record)

    """Supplies a dictionary of the errors in the sheet so the errors can be highlighted using openpyxl
    Returns:
        Dict: Dictionary of the errors in the sheet
    """
    def getSheetErrorDict(self):
        errorDict = {}
        for file in self.fileList:
            for key in file.errors:
                if file.errors[key]:
                    errorDict[file.fileName] = key
        return errorDict
    
    def getSheetFailureDict(self):
        failureDict ={}
        for file in self.fileList:
            for key in file.failures:
                if file.failures[key]:
                    failureDict[file.fileName] = key
        return failureDict
    



class ExcelFile():

    # Colors
    # Incorrect Filename - Orange
    # Duplicate Filename - Blue
    # Incorrect Date - Yellow
    # Failure - Red

    colorPalette = COLOR_PALETTE
    errorColors = {"Filename":"", "DupFilename":"", "Date":""} # This is outside the __init__ method so it is shared between all instances of the class
    failColors = {"Extent":"", "Filesize":"", "Existence":""} #Done this way to allow expansion of the fail types
    cachedColors = {}

    def __init__(self, filepath):
        self.filePath = filepath
        self.sheetList: List[ExcelSheet] = list()
        self.dataFrames = None
        self.retrieveErrorColors()

    def reset_file_structure(self):
        self.sheetList: List[ExcelSheet] = list()
        self.dataFrames = None
        self.retrieveErrorColors()


    """Allows the user to change the spreadsheet in the UI without creating a new object instance
        Clears the old sheet list, which in turn clearsthe file storage
    Args:
        newpath (str): the new filepath selected
    """
    def setFilePath(self) -> bool:

        new_path = easygui.fileopenbox(
            default=os.path.join(os.path.expanduser("~"), "Desktop,", "")
        )
        if new_path:
            self.filePath = new_path
            self.sheetList: List[ExcelSheet] = []
            return True
        else:
            return False

    def getTotalError(self):
        return sum([sheet.errors for sheet in self.sheetList])
    
    def getTotalFailures(self):
        return sum([sheet.failures for sheet in self.sheetList])
    
    def getTotalFiles(self):
        return sum([len(sheet.fileList) for sheet in self.sheetList])
    
    def setErrorColor(self, errorType, newColor):
        self.errorColors[errorType] = newColor

    def setFailColor(self, failType, newColor):
        self.failColors[failType] = newColor

    def resetErrorColors(self):
        shutil.copyfile(DEFAULT_COLORS, COLOR_PALETTE)

    def retrieveErrorColors(self):
        with open(COLOR_PALETTE, 'r', newline='') as file:
            csv_reader = csv.reader(file)
            header = next(csv_reader)  # Skip the header row
            for row in csv_reader:
                if row[0] in self.failColors:
                    self.failColors[row[0]] = row[1]
                else:
                    self.errorColors[row[0]] = row[1]
    
    def retrieveColorCache(self):
        with open(CACHED_COLORS, "r") as file:
            line = file.readline().strip()  # Read the first line
            values = line.split(",")
            values.pop()  # Split by comma, remove empty space

        # Create a dictionary with sequential keys
        self.cachedColors = {f"c{i+1}": value for i, value in enumerate(values)}

    def extendColorCache(self, color):
        with open(CACHED_COLORS, 'a+') as file:
            file.write(color + ",")
            file.close()
        self.retrieveColorCache()

    def clearColorCache(self):
        with open(CACHED_COLORS, "w") as file:
            pass  # No content is written, file is cleared



    # Defaults
    # Filename,FFEFBE7D
    # DupFilename,FF8BD3E6
    # Date,FFE9EC6B
    # Extent,FFFFADB0
    # Filesize,FFFFADB0
    # Existence,FFFFADB0
    def writeErrorColors(self):
        with open(COLOR_PALETTE, 'w', newline='') as file:
            csv_writer = csv.writer(file)
            csv_writer.writerow(['ErrorType', 'Hex'])
            for error in self.errorColors.keys():
                csv_writer.writerow([error, self.errorColors[error]])
            for fail in self.failColors.keys():
                csv_writer.writerow([fail, self.failColors[fail]])

    """Creates the file structure of an excel file creating each sheet and adding it to the sheetlist
        Method calls sheet.createScanFileList to create the list of files from the sheet dataframe
    """
    # It might be a better design decision to put this in FileHandler but this works for now
    def createFileStructure(self):

        self.reset_file_structure()

        try:
            if self.filePath == None:
                raise Exception ("No Spreadsheet Selected")
            else:
                xl_file = pd.ExcelFile(self.filePath)
                self.dataFrames = pd.read_excel(xl_file, sheet_name=None)
                [self.sheetList.append(ExcelSheet(sheet)) for sheet in xl_file.sheet_names] #Creates a new sheet object from the sheetname
                [sheet.createScanFileList(self.dataFrames[sheet.sheetName]) for sheet in self.sheetList] #Creates the list of files from the sheet dataframe
        except Exception as error:
            return error
    
    """Supplies a dictionary of sheets with the errors nested inside
    Returns:
        Dict: Dictionary of the errors in the file
    """
    def getFileErrorDict(self):
        return {sheet.sheetName: sheet.getSheetErrorDict() for sheet in self.sheetList}
    
    
    """Method to add auto fail comments to the correct cells in the spreadsheet
        Updates the QC Results, QC Comments, QC initials columns depending on fail types"""
    # Can't test until I can access the onedrive file structure
    def updateDataFrames(self):
        for sheet in self.sheetList:
            sheetDict = sheet.getSheetFailureDict()
            for key in sheetDict:


                # Clear failures from dataframe before adding new ones
                # This means corrected autofails can be overwritten
                #if self.dataFrames[sheet.sheetName]['QC Results'] == "Fail" and self.dataFrames[sheet.sheetName]['QC Initials'] == "AUTO":
                #
                #    self.dataFrames[sheet.sheetName]['QC Results'] = ""
                #    self.dataFrames[sheet.sheetName]['QC Comments'] = ""
                #    self.dataFrames[sheet.sheetName]['QC Initials'] = ""

                try:
                    self.dataFrames[sheet.sheetName]['QC Results'] = self.dataFrames[sheet.sheetName]['QC Results'].astype(str).replace(to_replace='nan', value='', regex=True)
                    self.dataFrames[sheet.sheetName]['QC Comments'] = self.dataFrames[sheet.sheetName]['QC Comments'].astype(str).replace(to_replace='nan', value='', regex=True)
                    self.dataFrames[sheet.sheetName]['QC Initials'] = self.dataFrames[sheet.sheetName]['QC Initials'].astype(str) .replace(to_replace='nan', value='', regex=True)
                    # Order matters here - if an error is matched in the order; Existence, Filesize, Extent then the others are ignored
                    # This means the failures go in order of precedence
                    # A file may fail for more than one of these reasons, only the mopst important reason is recorded
                    if sheetDict[key] in ['Extent', 'Filesize', 'Existence']:
                        self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Results'] = 'Fail'
                        self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Initials'] = 'AUTO'
                    if sheetDict[key] == 'Existence':
                        self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Comments'] = 'File does not exist'
                    elif sheetDict[key] == 'Filesize':
                        self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Comments'] = 'Filesize too large'
                    elif sheetDict[key] == 'Extent':
                        self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Comments'] = 'Incorrect page count'
                except KeyError:
                    return KeyError
