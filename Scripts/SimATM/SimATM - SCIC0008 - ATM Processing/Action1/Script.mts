﻿Dim dt_TCID, dt_ScenarioName, dt_TestCase, dt_ExpectedResult

Call spLoadLibrary()
Call spInitiateData("Excel_Report.xlsx", "SimATM - SCIC0008 - ATM Processing.xlsx", "ICON0008-001")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_ScenarioName, dt_TestCase, dt_ExpectedResult, "")

Call DA_Login_Bitvise("FNSTPCA")
Call Open_Terminal_Bitvise()
Call RUN_Command_Bitvise("SIMATM")
Call RUN_Command_Bitvise("SHOWJNL")
Call GET_JLN_Number()
Call GET_Fm_Acc_Number()
Call Close_Terminal_Bitvise()
Call DA_Logout_Bitvise()
Call DA_Login_Bitvise("FNSONLAD")
Call Open_Terminal_Bitvise()
Call RUN_Command_Bitvise("GETGLIF")
Call Close_Terminal_Bitvise()
Call DA_Logout_Bitvise()
Call DA_Login_ICONS()
Call GoTo_ScreenNumber()
Call Search_Inquiry_Trx_450()
Call DA_Logout_ICONS_No_SS()

Call spReportSave()

REM ========== SUB LOAD LIBRARY
Sub spLoadLibrary()
	Dim objSysInfo, Path_Env, LibFunction
	Set objSysInfo 		= Createobject("Wscript.Network")	
	Path_Env = Environment.Value("Path_Folder")
	LibFunction = Path_Env & "Libraries\"
	LibRepo = Path_Env & "Repositories\"
	LoadFunctionLibrary (LibFunction & "Lib_Report.vbs")
	LoadFunctionLibrary (LibFunction & "Lib_GlobalFunction.qfl")
	
	LoadFunctionLibrary (LibFunction & "Lib_BitviseLog.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_BitviseTerminal.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_Notepad.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_ICONS.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_Screen_450.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_ICONS_Main.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Notepad.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Bitvice.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Screen_450.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Bitvice_Terminal.tsr")
End Sub

Sub spGetDatatable()
	REM ---------- Report Data
	dt_TCID					= DataTable.Value("TC_ID", dtLocalSheet)
	dt_ScenarioName			= DataTable.Value("SCENARIO_NAME", dtLocalSheet)
	dt_TestCase				= DataTable.Value("TEST_CASE", dtLocalSheet)
	dt_ExpectedResult		= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
End Sub

