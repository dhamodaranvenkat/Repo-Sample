'strPatN = "aut1,Tes"
'strFacil = "TriHealth Hospital"
'If Window("EpicWindow-PT").InsightObject("LandingPage_btn_Maximize").Exist (3) Then	
'	Window("EpicWindow-PT").InsightObject("LandingPage_btn_Maximize").Click
'End If
'
'
'Window("EpicWindow-PT").InsightObject("PatLkpUp_Edit_Search").Type strPatN
'Window("EpicWindow-PT").Type micTab
'Set objPatLkpList = Window("EpicWindow-PT").InsightObject("PatLkpUp_List_Click")
'
'fn_RowSelect Cint(442), Cint(750), Cint(-45), Cint(45), Cint(16), Cint(16), objPatLkpList, strPatN
'
'Set objPatList = Window("EpicWindow-PT").InsightObject("PatAvailableList_Lst_Info")
'
'fn_RowSelect Cint(90), Cint(750), Cint(-80), Cint(150), Cint(10), Cint(15), objPatList, strFacil
'
'
'
'Window("EpicWindow-PT").InsightObject("PatLkpUp_Edit_Search").Type strPatN
''fn_SendKeys Ucase("a")
'fn_SendKeys "Enter"
'Set objPatLkpList = Window("EpicWindow-PT").InsightObject("PatLkpUp_List_Click")
'fn_RowSelect Cint(442), Cint(750), Cint(-45), Cint(45), Cint(16), Cint(16), objPatLkpList, strPatN
'Set objPatList = Nothing
'
'Public Function fn_SendKeys(strKeyName)
'	Set oshell = CreateObject("Wscript.Shell")
'	Select Case Ucase(strKeyName)
'		Case "ENTER"		
'    		oshell.SendKeys "{ENTER}"
'		Case "TAB"
'    		oshell.SendKeys "{TAB}"
'		Case "A"
'    		oshell.SendKeys "{A}"
'	End Select
'    Set oshell = Nothing
'End Function
'


'Reporter.Filter = 3
intAlmId = inputbox ("Enter the AlmId, Only integer Between 1 to 4")

	gTotalDataSheet = Datatable.GetSheetCount
	'TestCase Details Screen Input
	For iTestCaseDetails = 1 to Datatable.GetSheet("TESTCASE_DETAILS").GetRowCount
		Datatable.GetSheet("TESTCASE_DETAILS").SetCurrentRow(iTestCaseDetails)
		If Datatable("Alm_ID","TESTCASE_DETAILS") = intAlmId Then
			'Need to check		
			Datatable.GetSheet("TESTCASE_DETAILS").SetCurrentRow (iTestCaseDetails)
			strPatientName = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("Patient_Name"))
			strPatientPhoneNum = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("Patient_Phone_Number"))
			strPatientEmergencyCont = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("Emergency_Contact"))
			strPatientTypeOfAdmit = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("TypeOfAdmit"))
			strPatientDOB = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("DOB"))
			strPatientSex = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("SEX"))
			strPatientAddress = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("PatientAddress"))
			strPatientCtySteZip = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("CityStateZip"))
			strPatientAge = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("Age"))
			strPatientRelationship = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("PatientRelationship"))
			strUserID = Trim(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("UserID"))
			strEnv = Trim(Ucase(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("Env")))
			strAdmissionType = Trim(Ucase(Datatable.GetSheet ("TESTCASE_DETAILS").GetParameter("AdmissionType")))
	Exit For	
		End If
			Datatable.GetSheet("TESTCASE_DETAILS").SetNextRow
	Next
	
	'UserID screen Input
	For iUserID = 1 To Datatable.GetSheet("USER_ID").GetRowCount
		Datatable.GetSheet("USER_ID").SetCurrentRow(iUserID)	
		If Datatable("UserName","USER_ID") = strUserID Then
			Datatable.GetSheet ("USER_ID").SetCurrentRow (iUserID)			
			strUserName = Datatable.GetSheet ("USER_ID").GetParameter("UserName")
			strPassword = Datatable.GetSheet ("USER_ID").GetParameter("Password")
	Exit For
		End If
		Datatable.GetSheet("USER_ID").SetNextRow
	Next
	
	'For BedSelection Screen
	For iBedSelection = 1 To Datatable.GetSheet("BED_SELECTION").GetRowCount
		Datatable.GetSheet("BED_SELECTION").SetCurrentRow(iBedSelection)	
		If Datatable("UserName","BED_SELECTION") = strUserID Then
			Datatable.GetSheet ("BED_SELECTION").SetCurrentRow (iBedSelection)
			If strEnv = "TST" Then
				intPattern = Datatable.GetSheet ("BED_SELECTION").GetParameter("TST_Pattern")
			else
				intPattern = Datatable.GetSheet ("BED_SELECTION").GetParameter("POC_Pattern")
			End If
			intBedPosition = Datatable.GetSheet ("BED_SELECTION").GetParameter("BedPosition")
			strBedAdmissiontype = Datatable.GetSheet ("BED_SELECTION").GetParameter("AdmissionType")
	Exit For
		End If
		Datatable.GetSheet("BED_SELECTION").SetNextRow
	Next
	
	'fn_EnvironmentSelection strEnv
	fn_Login strUserName, strPassword, strAdmissionType
	intPatientNum = fn_SelectPatientBeforeBed(strPatientName)
	fn_PatientDragDropOpenBed intPatientNum, intBedPosition, strBedAdmissiontype
	fn_AcceptBed strPatientName
	fn_PendingSummaryClick (strPatientName)
	fn_Demographics strPatientPhoneNum, strPatientRelationship, strPatientEmergencyCont, strPatientCtySteZip, strPatientAddress,strPatientAge , intAlmId
	fn_Logout
