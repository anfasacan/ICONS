
Dim dt_TCID, dt_ScenarioName, dt_TestCase, dt_ExpectedResult


Call spLoadLibrary()
Call spInitiateData("Excel_Report.xlsx", "SCIC0004 - Testing Mandatory Field.xlsm", "sheetname")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_ScenarioName, dt_TestCase, dt_ExpectedResult, "")
Iteration = Environment.Value("ActionIteration")

Call DA_Login_ICONS()
Call GoTo_ScreenNumber()


If Iteration = 1 Or Iteration = 3  Then
	Call Filling_Detail_CIF_63001
	Call Filling_Kontak_CIF_63001
	Call Generate_CIF_63001()

ElseIf Iteration = 5 Then
	Call Filling_Field_Detail_63001_Perusahaan()
	Call Filling_Field_Kontak_63001_Perusahaan()
	Call Generate_CIF_63001()
	
ElseIf Iteration = 7 Then
	Call Filling_Detail_CIF_63001
	Call Filling_Kontak_CIF_63001
	Call Generate_CIF_63001()
	
ElseIf Iteration = 2 Then
	Call Check_Mandatory_Field_63001_Perorangan()
	
ElseIf Iteration = 4 or Iteration = 6 Then
	Call Check_Mandatory_Field_63001_Perusahaan()

ElseIf Iteration = 8 Then
	Call Check_Mandatory_Field_63001_Pemerintahan()
	
'ElseIf Iteration > 2 and Iteration < 34 Then
'	Call Cleaning_Field_63001_Perorangan()
'	Call Filling_Field_Detail_63001_Perorangan()
'	Call Cleaning_Auto_Fill_Field_63001_Perorangan()
'	Call Check_Mandatory_Field_is_True_36001_Detail_Perorangan()
'
'ElseIf Iteration > 33 Then
'	Call Filling_Field_Detail_63001_Perorangan()
'	Call Filling_Field_Kontak_63001_Perorangan()
'	Call Check_Mandatory_Field_is_True_36001_Kontak_Perorangan()
End If

Call DA_Logout_ICONS()

	
Call spReportSave() @@ hightlight id_;_Browser("ICONS 64 Alternity").Page("ICONS 64 Alternity").Frame("MainScreen").WebList("notipros")_;_script infofile_;_ZIP::ssf25.xml_;_


REM ========== SUB LOAD LIBRARY
Sub spLoadLibrary()
	Dim objSysInfo, Path_Env, LibFunction
	Set objSysInfo 		= Createobject("Wscript.Network")	
	Path_Env = Environment.Value("Path_Folder")
	LibFunction = Path_Env & "Libraries\"
	LibRepo = Path_Env & "Repositories\"
	LibExcel = Path_Env & "Excel\"
	LoadFunctionLibrary (LibFunction & "Lib_Report.vbs")
	LoadFunctionLibrary (LibFunction & "Lib_GlobalFunction.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_BitviseTerminal.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_BitviseLog.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_Screen_63001.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_ICONS.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_ICONS_Main.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Bitvice.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Screen_63001.tsr")
	
End Sub

Sub spGetDatatable()
	REM ---------- Report Data
	dt_TCID					= DataTable.Value("TC_ID", dtLocalSheet)
	dt_ScenarioName			= DataTable.Value("SCENARIO_NAME", dtLocalSheet)
	dt_TestCase				= DataTable.Value("TEST_CASE", dtLocalSheet)
	dt_ExpectedResult		= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
End Sub

