#Author-Patrick Owens
#Description-Reads in a CSV file where the first column is an ID value and the remaining columns are user parameters.\nFor each row in the file the parametric values will be applied to the model and the resulting model output as a STEP file.\nColumn headings that don't match a user parameter are skipped, allowing the file to also contain other values for CFD...

import adsk.core, adsk.fusion, traceback
from pathlib import Path
from csv import DictReader

def run(context):
    ui = None

    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        design = app.activeProduct
        rowCount = 0
        rowOutput = 0

        # Set styles of CSV file dialog
        fileDlg = ui.createFileDialog()
        fileDlg.isMultiSelectEnabled = False
        fileDlg.title = 'Select a .csv Configuration File'
        fileDlg.filter = '*.csv'

        # Show file open dialog
        dlgResult = fileDlg.showOpen()
        # If the selection is OK
        if dlgResult == adsk.core.DialogResults.DialogOK:
            # Open the CSV file:
            with open(fileDlg.filename, 'r') as csvFile:
                csvDict = DictReader(csvFile)
                # Process the row
                for row in csvDict:
                    rowCount += 1
                    # Set all the parameter values for the row
                    paramResults = setParameters(design, row)
                    # Verify the parameters were set
                    if paramResults == True: # all were set without issue
                        # get the value of the first field in the CSV and build the full path for the output file
                        outputFile = str(Path(fileDlg.filename).parent.joinpath(list(row.values())[0] + '.step'))
                        print(outputFile)
                        # call exportSTEP() to output the file
                        fileResult = exportSTEP(design, outputFile)
                        if fileResult == True:
                            rowOutput += 1
                        else:
                            # Warn the user that the file was not written
                            print(fileResult)
                            fileErrorChoice = ui.messageBox(
                                fileResult + '\nPress OK to continue processing CSV file: ', 
                                'Bulk STEP Export', 
                                adsk.core.MessageBoxButtonTypes.OKCancelButtonType, 
                                adsk.core.MessageBoxIconTypes.WarningIconType)
                            
                            if fileErrorChoice == adsk.core.DialogResults.DialogOK: # DialogOK = 0
                                # Error writing file, user wants to process the next row
                                print('OK, move on to next row')
                                continue # skip this file but process the next row
                            elif paramErrorChoice == adsk.core.DialogResults.DialogCancel: # DialogCancel = 1
                                # Error writing file, user wants to exit script
                                print('cancel out of CSV loop')
                                break
                            else:
                                # give up and exit
                                print('This is not my beautiful wife!')
                                break
                    else: # paramResults != True
                        # if setting the keys didn't work, need to ask if we want to cancle the whole thing or just this row
                        paramErrorChoice = ui.messageBox(
                            paramResults + '\nPress OK to continue processing csv file', 
                            'Bulk STEP Export', 
                            adsk.core.MessageBoxButtonTypes.OKCancelButtonType, 
                            adsk.core.MessageBoxIconTypes.WarningIconType)
                        
                        if paramErrorChoice == adsk.core.DialogResults.DialogOK: # DialogOK = 0
                            # Error on this row but user wants to process the next
                            print('OK, move on to next row')
                            continue # skip this file but process the next row
                        elif paramErrorChoice == adsk.core.DialogResults.DialogCancel: # DialogCancel = 1
                            # Error on this row so exit program
                            print('cancel out of CSV loop')
                            break
                        else:
                            # give up and exit
                            print('This is not my beautiful house!')
                            break
                    # to refresh the ui
                    #returnValue = command_var.doExecute(terminate)
        # If something went wrong or no file was selected
        else:
            ui.messageBox(
                'Quit: No File Selected', 
                'Bulk STEP Export', 
                adsk.core.MessageBoxButtonTypes.OKButtonType, 
                adsk.core.MessageBoxIconTypes.CriticalIconType)
        ui.messageBox(
            'Done Processing {} rows in CSV\nOutput {} files'.format(rowCount, rowOutput), 
            'Bulk STEP Export', 
            adsk.core.MessageBoxButtonTypes.OKButtonType, 
            adsk.core.MessageBoxIconTypes.NoIconIconType)
        print('exit script')
        return
        
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()), 'Bulk STEP Export')
        return

def setParameters(design, row):
    key = ''
    value = ''
    try:
        for key, value in row.items():
            # check each key and set its value
            param = design.userParameters.itemByName(key)
            if param:
                print(key, value)
                param.expression = value
        return True
    except Exception as e:
        print(str(e))
        return 'Error: {}\nWhile setting parameter {} to {}\n'.format(e, key, value)

def exportSTEP(design, path):
    fileResult = None
    #Export the STEP file.
    try:
        exportMgr = design.exportManager
        stepOptions = exportMgr.createSTEPExportOptions(path)
        return exportMgr.execute(stepOptions) # returns true if successful
    except Exception as e:
        print(str(e))
        return 'Error: {}\nWhile exporting STEP file: {}\n'.format(e, path)
