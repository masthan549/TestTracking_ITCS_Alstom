import re, os, sys, xlrd, xlsxwriter, time
from importlib import reload
from datetime import datetime 
from tkinter import messagebox

worksheet_html = ""
worksheet_html_start_rowCounter = 0
workbook = ""
testResPath = ""
testSeqPath = ""

softwareVer = ""
TestLabelVer = ""
testSheetName = ""

# HTML REPORT: Read sequence and test case and test result.
def getTestCaseStatus(seqName, DictOfTCs):

    global testResPath
    global testSeqPath
   
	#Fill the status of every sequence testcase in format like = 
	#[{testcaseNumber1:PASS/FAIL/NA/ERROR, testcaseNumber2:PASS/FAIL/NA/ERROR}, {testcaseNumber2:PASS/FAIL/NA/ERROR, testcaseNumber2:PASS/FAIL/NA/ERROR}]
    SeqAndTCStatus = DictOfTCs
    seqName = seqName+"_Report"

    HTMLFileNamesInDir = os.listdir(testResPath)
    AvailableReportsForSeq = [seqReportName for seqReportName in HTMLFileNamesInDir if seqName.lower() in seqReportName.lower()]

    #This needs to be repeated for multiple tests.

    for file in AvailableReportsForSeq:
        if os.path.isfile(testResPath+file) and file.endswith(".html"):
		
			#Open sequence file and read for testcase numbers
            fPtr = open(testResPath+file)
            line = fPtr.readline()
    		
            global lookForTestcaseStatus		
            lookForTestcaseStatus = False		
            while line:
                regData = re.search(r".*<TD.*>(OBCTP\d*.*_TC.*) </TD>.*",line,re.I)
                try:
                    testcaseName = (regData.groups(0)[0]).strip()
                    lookForTestcaseStatus = True
                except:pass
                line = fPtr.readline()
                    
                #Look for testcase status
                while lookForTestcaseStatus and line:			
                    regData_status = re.search(r".*<FONT.*>(.*)<.*FONT>.*",line,re.I)
                    try:
                        tcStatus = regData_status.groups(0)[0]
                        if ((tcStatus == "Passed") or (tcStatus == "Failed") or (tcStatus == "Skipped")):
    						
                            if (tcStatus == "Passed"):
                                SeqAndTCStatus[testcaseName] = "PASS"
                            elif (tcStatus == "Failed"):
                                SeqAndTCStatus[testcaseName] = "FAIL"
                            elif (tcStatus == "Skipped"):
                                SeqAndTCStatus[testcaseName] = "SKIP"
								
                            lookForTestcaseStatus = False
                            line = fPtr.readline()
                    except: pass
    				
                    if lookForTestcaseStatus is True:			
                        line = fPtr.readline()

    return SeqAndTCStatus 

	
def writeTCAndStatusIntoSheet(seqName, testcaseStatus, TPTitle):

	
    global worksheet_html
    global worksheet_html_start_rowCounter
    global workbook
    
    
    # Colors	
    red = workbook.add_format({'bg_color': 'red', 'bold': 1, "font_name": "Alstom"})
    green = workbook.add_format({'bg_color': 'green', 'bold': 1, "font_name": "Alstom"})	
    yellow = workbook.add_format({'bg_color': 'yellow', 'bold': 1, "font_name": "Alstom"})	
    orange = workbook.add_format({'bg_color': 'orange', 'bold': 1, "font_name": "Alstom"})	
    AlstomFont = workbook.add_format({'align': 'left', "font_name": "Alstom", 'size': 11})	
    
    #Write each testcase status into sheet
    for dictIndx in sorted(testcaseStatus):
        worksheet_html.write('C'+str(worksheet_html_start_rowCounter),dictIndx, AlstomFont)	
        if (testcaseStatus[dictIndx] == "PASS"):						
            worksheet_html.write('D'+str(worksheet_html_start_rowCounter),"PASS",green)	
        elif (testcaseStatus[dictIndx] == "FAIL"):
            worksheet_html.write('D'+str(worksheet_html_start_rowCounter),"FAIL",red)	
        elif (testcaseStatus[dictIndx] == "SKIP"):
            worksheet_html.write('D'+str(worksheet_html_start_rowCounter),"SKIP",orange)				
        elif (testcaseStatus[dictIndx] == "NA"):
            worksheet_html.write('D'+str(worksheet_html_start_rowCounter),"RESULTS NOT AVAILABLE",yellow)
			
        worksheet_html.write('B'+str(worksheet_html_start_rowCounter),seqName, AlstomFont)
        worksheet_html.write('A'+str(worksheet_html_start_rowCounter),TPTitle, AlstomFont)
        worksheet_html_start_rowCounter = worksheet_html_start_rowCounter + 1



def writeExcelHeader():

	
    global worksheet_html
    global worksheet_html_start_rowCounter
    global workbook, testSheetName
	
    dt = datetime.now()
    day_hour_min_sec = str(dt.day)+"_"+str(dt.hour)+"_"+str(dt.minute)+"_"+str(dt.second)
    testSheetName = 'TestResultSheet'+'_'+str(day_hour_min_sec)+'.xlsx'	
	
    # Creating Xlsheet to write the data into that if it is not exist 	
    try:
        workbook = xlsxwriter.Workbook('TestResultSheet'+'_'+str(day_hour_min_sec)+'.xlsx')
    except:
        #messagebox.showerror('Error','Please close TestResultSheet xlsx file if it is opened')          
        print("Please close the sheet and try again")
        sys.exit()
    
    # Colors	
    bold = workbook.add_format({'bg_color': '#A6A6A6', 'bold': 1, 'align': 'left', "font_name": "Alstom", 'size': 12})
   
    # Write the sequence and testcase and its result into workbook sheet
    worksheet_html = workbook.add_worksheet('OBC SW Build '+str(softwareVer))
    worksheet_html_start_rowCounter = 12
    worksheet_html.write('A11',"Procedure", bold)
    worksheet_html.write('B11',"Script", bold)
    worksheet_html.write('C11',"Test Case", bold)
    worksheet_html.write('D11',"Results", bold)
    worksheet_html.write('E11',"CR", bold)
    worksheet_html.write('F11',"CR Type", bold)
    worksheet_html.write('G11',"CR Status", bold)
    worksheet_html.write('H11',"Comments / Remarks ", bold)
    worksheet_html.write('I11',"First Validator Validator ", bold)
    worksheet_html.write('J11',"Second Validation", bold)

    # Set width of the columns
    worksheet_html.set_column(0,0,43)	
    worksheet_html.set_column(1,1,20)	
    worksheet_html.set_column(2,2,20)	
    worksheet_html.set_column(3,3,23)	
    worksheet_html.set_column(7,7,23)
    worksheet_html.set_column(8,8,23)	
    worksheet_html.set_column(9,9,23)
    worksheet_html.set_row(2,45)	
	

def updateSweepMetricsHeader():
	
    global worksheet_html
    global workbook
	
    # Colors	
    bold = workbook.add_format({'bold': 1, 'align': 'center', 'size': 18, "font_name": "Alstom"})
	

    merge_format = workbook.add_format({'bg_color': '#DDD9C4', 'bold': 1, 'align': 'center', 'valign': 'vcenter', 'size': 14, "font_name": "Alstom"})
	
    bold_header = workbook.add_format({'bg_color': '#A6A6A6', 'bold': 1, 'align': 'center', 'size': 12, "font_name": "Alstom"})	
    bold_percentage_format = workbook.add_format({'bold': 1, 'num_format': '0.00%', "font_name": "Alstom"})	
    bold_size_format = workbook.add_format({'bold': 1, 'size': 10, "font_name": "Alstom"})	
	
    bold_green_format = workbook.add_format({'bg_color': '#00B050', 'bold': 1, 'size': 10, "font_name": "Alstom"})	
    bold_red_format = workbook.add_format({'bg_color': '#C00000', 'bold': 1, 'size': 10, "font_name": "Alstom"})	
    bold_yellow_format = workbook.add_format({'bg_color': 'yellow', 'bold': 1, 'size': 10, "font_name": "Alstom"})	
    bold_blue_format = workbook.add_format({'bg_color': '#0070C0', 'bold': 1, 'size': 10, "font_name": "Alstom"})	
	
    bold_green_percentage_format = workbook.add_format({'bg_color': '#00B050', 'bold': 1, 'size': 10, 'num_format': '0.00%', "font_name": "Alstom"})	
    bold_red_percentage_format = workbook.add_format({'bg_color': '#C00000', 'bold': 1, 'size': 10, 'num_format': '0.00%', "font_name": "Alstom"})	
    bold_yellow_percentage_format = workbook.add_format({'bg_color': 'yellow', 'bold': 1, 'size': 10, 'num_format': '0.00%', "font_name": "Alstom"})	
    bold_blue_percentage_format = workbook.add_format({'bg_color': '#0070C0', 'bold': 1, 'size': 10, 'num_format': '0.00%', "font_name": "Alstom"})		
	
    # Write sweep metrics header	
    worksheet_html.write('A1',"Sweep Metrics "+str(softwareVer), bold)	
    worksheet_html.merge_range('A3:D3', 'Validation Label: ITCS Validation Version '+str(TestLabelVer), merge_format)	
	
    worksheet_html.write('A4',"DESCRIPTION", bold_header)	
    worksheet_html.write('B4',"UOM", bold_header)	
    worksheet_html.write('C4',"Total OBC", bold_header)	
    worksheet_html.write('D4',"% OBC", bold_header)	
	
    worksheet_html.write('A5',"TOTAL TEST CASES", bold_size_format)	
    worksheet_html.write('B5',"#", bold_size_format)	
    worksheet_html.write('C5',"=(COUNTA(C12:C"+str(worksheet_html_start_rowCounter-1)+"))", bold_size_format)
    worksheet_html.write('D5',"=SUM(D6:D9)", bold_percentage_format)	

    worksheet_html.write('A6',"TC PASSED", bold_size_format)	
    worksheet_html.write('B6',"#", bold_size_format)	
    worksheet_html.write('C6',"=(COUNTIF(D12:D"+str(worksheet_html_start_rowCounter-1)+", \"PASS\"))", bold_green_format)
    worksheet_html.write('D6',"=(C6/(C6+C7+C8+C9))", bold_green_percentage_format)	

    worksheet_html.write('A7',"TC CONDITIONALLY PASSED", bold_size_format)	
    worksheet_html.write('B7',"#", bold_size_format)	
    worksheet_html.write('C7',"=(COUNTIF(D12:D"+str(worksheet_html_start_rowCounter-1)+", \"C-PASS\"))", bold_yellow_format)
    worksheet_html.write('D7',"=(C7/(C6+C7+C8+C9))", bold_yellow_percentage_format)
	
    worksheet_html.write('A8',"TC FAILED", bold_size_format)	
    worksheet_html.write('B8',"#", bold_size_format)	
    worksheet_html.write('C8',"=(COUNTIF(D12:D"+str(worksheet_html_start_rowCounter-1)+", \"FAIL\"))", bold_red_format)
    worksheet_html.write('D8',"=(C8/(C6+C7+C8+C9))", bold_red_percentage_format)	
	
    worksheet_html.write('A9',"TC SKIPPED", bold_size_format)	
    worksheet_html.write('B9',"#", bold_size_format)	
    worksheet_html.write('C9',"=(COUNTIF(D12:D"+str(worksheet_html_start_rowCounter-1)+", \"SKIP\"))", bold_blue_format)
    worksheet_html.write('D9',"=(C9/(C6+C7+C8+C9))", bold_blue_percentage_format)	

    # add borders to sheet
    border_format=workbook.add_format({'border':1, "font_name": "Alstom"})	
    worksheet_html.conditional_format('A1:J'+str(worksheet_html_start_rowCounter-1),{ 'type' : 'no_blanks' , 'format' : border_format})	
	
def checkIfTestCasesInSeqSkipped(TCStatus):

    testcaseExecuted = False
    testDict = TCStatus
   
    # Mark testcase status as skipped if atleast one testcase is PASS/FAIL/SKIP
    for tcStatusIndx in testDict:
        if ((testDict[tcStatusIndx] == "PASS") or (testDict[tcStatusIndx] == "FAIL") or (testDict[tcStatusIndx] == "SKIP")):
            testcaseExecuted = True
            break
			
    # Update the status if testcaseExecuted is True
    if testcaseExecuted:
        for tcStatusIndx in testDict:
            if (testDict[tcStatusIndx] == "NA"):
                testDict[tcStatusIndx] = "SKIP"
    return testDict

def compareResultsWithPreviousBuildTestTrackingSheet(currentBuildTestTrackingSheet, prevBuildTestTRackingWorkbook_arg, prevBuildTestTrackingSheet_arg):

    pass
	
#if __name__ == "__main__":

def script_exe(testSeqPath_arg, testResPath_arg, buildVer, labelVer, TkObject_ref, statusBarText):

    #testSeqPath = "C:\\Results Analysis\\Sequence\\test\\"
    #testResPath = "C:\\Results Analysis\\Results\\"	

    global testSeqPath, softwareVer
    global testResPath, TestLabelVer
    global TkObject,statusBar	

    testSeqPath = testSeqPath_arg+"/"
    testResPath = testResPath_arg+"/"
    statusBar = statusBarText	
    TkObject = TkObject_ref	

	
    softwareVer = buildVer
    TestLabelVer = labelVer	
	
    #Write Excel header
    writeExcelHeader()
    dirList = sorted(os.listdir(testSeqPath))

    lstFile = [f for f in os.listdir(testSeqPath) if f.endswith('.seq')]
    progressCounter = 0
    
    # SEQUENCE: Read list of sequences in directory given in testSeqPath
    for file in dirList:
        if os.path.isfile(testSeqPath+file) and file.endswith(".seq"):
		
            statusBar.set("Number of sequences completed: ("+str(progressCounter)+"/"+str(len(lstFile))+")."+" Currently Analysing sequence: \""+ file+"\". Please wait...")
            progressCounter = progressCounter+1
		
    		#Fill the status of every sequence testcase in format like = 
    		#{testcaseNumber1:PASS/FAIL/NA/ERROR, testcaseNumber2:PASS/FAIL/NA/ERROR}
            FinalTestStatus = {}
    		
            seqName = file[:-4]
            DictOfTCInSeq = {}
            TPTitle = "NA"
    
    	
            #Open sequence file and read for testcase numbers
            #fPtr = open(testSeqPath+file,'r',encoding="cp1252")
            fPtr = open(testSeqPath+file,'r',encoding="Latin-1")
            line = fPtr.readline()		
            while line:
			
                # Fecth testcase name
                regData = re.search(r"SeqName.*(OBC[TP]{0,1}\d*.*TC.*$)",line,re.I)
                try:		
                    testcaseName = regData.groups(0)[0]
                    testcaseName = testcaseName[:-1]					
                    DictOfTCInSeq[testcaseName] = "NA"
                except:pass

                # Fecth sequence name
                regData_TPName = re.search(r"TPTitle.*=.*\"(.*)\"",line,re.I)
                try:		
                    TPTitle = regData_TPName.groups(0)[0]
                except:pass
				
                #While reading line by line text from files if compiler sees any charector which is not understandble, then ignore such lines
                while True:			
                    try:
                        line = fPtr.readline()
                        break
                    except: 
                        pass			

						
            #Get each testcase status from HTML report and update the status in dictionary
            FinalTestStatus = getTestCaseStatus(seqName,DictOfTCInSeq)
			
			
            #See if any testcases in sequence files are skipped, if so then update the testcase status as SKIP 
            FinalTestStatus_updated = checkIfTestCasesInSeqSkipped(FinalTestStatus)

            
            #Write test result into sheet
            writeTCAndStatusIntoSheet(file, FinalTestStatus_updated, TPTitle)

    updateSweepMetricsHeader()			
    workbook.close()
	
	
    #Compare results with previous builds
    #prevBuildTestTRackingWorkbook = "C:\Results Analysis\OBC Test Tracking (Builds 0800 - Present).xlsx"
    #prevBuildTestTRackingSheet = "OBC SW Build 0926.0A"
    #
    #compareResultsWithPreviousBuildTestTrackingSheet(str(os.getcwd())+"\\"+testSheetName, prevBuildTestTRackingWorkbook, prevBuildTestTRackingSheet)	
	
    statusBar.set("DONE!!")	
	
	
    messagebox.showinfo('DONE!!',"Test tracking sheet with results produced in : "+str(os.getcwd())+"\\"+testSheetName)
    TkObject.destroy()
    sys.exit()	

