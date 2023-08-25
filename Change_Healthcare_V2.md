Option Compare Database


'Delete Button
Private Sub cmdDelete_Click()
	If MsgBox("Are you want to delete old data?", vbYesNo) = vbYes Then
		DoCmd.SetWarnings (False)
		DoCmd.RunSQL ("delete * from tbl_InputData")
		MsgBox " Data Deleted Successfully!"
	End If
End Sub

Private Sub cmdIE_Edge_Click()
	driver.Start "Edge"
	driver.Get "http://access.emdeon.com"
End Sub

'Import Button
Private Sub cmdImport_Click()
	
'DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl_InputData", CurrentProject.Path & "\Tb1_Inputdata_New.xlsx", True
	DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl_InputData", CurrentProject.Path & "\Tbl_Inputdata_New.xlsx", True
	DoCmd.RunSQL ("Delete from tbl_InputData where [rx nbr] is null")
'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "tbl_InputData", CurrentProject.Path & "\Output1.xlsx"
'    DoCmd.SetWarnings False
'    Module2.Select_file
'
'    If strpath <> "" Then
'        DoCmd.RunSQL " delete * from tbl_Congnos"
'        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tbl_Congnos", strpath, True
	MsgBox "Uploaded"
'    Else
'        MsgBox "Please select file"
'    End If
'
'        DoCmd.SetWarnings True
End Sub
'Login into Website using given creditionals
Private Sub cmdLog_Click()
'Dim myIE As New InternetExplorer
'DoCmd.SetWarnings (False)
'myPageURL = "https://apps.availity.com/"
	
'    driver.Start "Edge"
'
'    driver.Get "https://apps.availity.com/"
'    driver.SwitchToDefaultContent
'
	
'Set myIE = GetNewPage
' myIE.navigate myPageURL
' myIE.Visible = True
' myIE.FullScreen = False
' Do While myIE.Busy = True Or myIE.ReadyState <> 4
'    DoEvents
'Loop
	
	
' myIE.Document.getElementById("userId").Value = txtUser
' driver.FindElementById("userId").Value = txtUser
	
' myIE.Document.getElementById("password").Value = txtPassword
' driver.FindElementById("password").Value = txtPassword
	
' myIE.Document.getElementById("loginFormSubmit").Click
'driver.FindElementById("loginFormSubmit").Click
	
' Do While myIE.Busy = True Or myIE.ReadyState <> 4
'    DoEvents
'Loop
	
'Sleep (5000)
	
'    For Each objInputElement In driver.FindElementsByTag("input")
'        If objInputElement.Type = "button" And objInputElement.Text = "Continue" Then
'            objInputElement.Click
': Exit For
'        End If
'    Next objInputElement
'
'    Do While myIE.Busy = True Or myIE.ReadyState <> 4
'        DoEvents
'    Loop
'
'    Me.Dirty = False
'    Do While myIE.Busy = True Or myIE.ReadyState <> 4
'        DoEvents
'    Loop
'
'    Do While myIE.Busy = True Or myIE.ReadyState <> 4
'        DoEvents
'        Loop
'        Sleep (10000)
'MsgBox ("Login Process Completed")
'
'DoCmd.SetWarnings (True)
	
End Sub
'Get Data from website
Private Sub cmdLogin_Click()
	If MsgBox("Have you selected STATE as Illinois on Website?", vbYesNo, "Question") = vbYes Then
		GotoAvailitySiteNew
	End If
End Sub

Private Sub cmdPatient_Search_Click()
	
	Dim str As String
	Dim clms As String
	Dim rs As New ADODB.Recordset
	rs.Open "Select * from tbl_InputData where [CCN/ICN] is null ", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
	While Not rs.EOF
		
		StartDate = CDate(rs![Fill Entered Dttm]) - 15
		EndDate = CDate(rs![Fill Entered Dttm]) + 15
		sdate = StartDate
		edate = EndDate
		sFdt = Format(StartDate, "Mmm") & " " & Format(StartDate, "d") & ", " & Format(StartDate, "YYYY")
		sTdt = Format(EndDate, "Mmm") & " " & Format(EndDate, "d") & ", " & Format(EndDate, "YYYY")
		
		str = "https://access.emdeon.com/ProviderVision/provider/getJspxReport.jspx?id=PR_1001_Insured&event_name_id="
		str = str & "Patient Search&from=" & sFdt & "&to=" & sTdt & "&npi=&date.type=SERVICE_FROM&insured.id=" & rs![General Recipient Nbr]
		str = str & "&dateFrom=" & StartDate & "&dateTo=" & EndDate & "&is_from_search_page=yes&fromPage=search&searchType="
		str = str & "Patient Search&parentURL=/ProviderVision/provider/getIndexPage.jspx&random=1651002180232"
		
		driver.ExecuteScript "window.open(arguments[0])", str
		Sleep (2000)
		driver.Window.Maximize
		driver.SwitchToNextWindow
		driver.SwitchToDefaultContent
		
		Dim Found As Integer
		Found = 0
		i = 1
		For Each rw In driver.FindElementById("reportContentTable").FindElementsByTag("tr")
'For Each rw In driver.FindElementByClass("tab-header table-style").FindElementByTag("tr")
			Debug.Print rw.Text
			On Error Resume Next
			cln = Split(rw.Text, " ")
			Debug.Print cln
			If Err.Number <> 438 Then
				If UBound(cln) > 3 Then
					If InStr(1, cln(3), "-") > 0 Then
'                        If CLng(Trim(Left(cln(3), InStr(1, cln(3), "-") - 1))) Like "" & Trim(rs![rx nbr]) & "*" And Trim(cln(4)) = Format(rs![Fill Entered Dttm], "mm/dd/yyyy") Then 'And Trim(cln(4)) = Format(rs!WebSDL_SOLD, "mm/dd/yyyy")   CLng(Trim(Left(cln(4), InStr(1, cln(4), "-") - 1)))
'                            clms = driver.FindElementByXPath("//*[@id=""reportListItemID" & i & """]/td[6]").Text
'                            driver.FindElementByXPath("//*[@id=""reportListItemID" & i & """]/td[8]/a").Click
'                            Sleep 2000
'                            Found = 1
'                            Exit For
'                        End If
'                    Else
						If Trim(cln(3)) Like "" & Trim(rs![rx nbr]) & "*" And Trim(cln(4)) = Format(rs![Fill Entered Dttm], "mm/dd/yyyy") Then 'Format(rs!WebSDL_SOLD, "mm/dd/yyyy")
							clms = driver.FindElementByXPath("//*[@id=""reportListItemID" & i & """]/td[6]").Text
							driver.FindElementByXPath("//*[@id=""reportListItemID" & i & """]/td[8]/a").Click
							Sleep 2000
							Found = 1
							Exit For
						End If
					End If 'instr
					i = i + 1
				End If 'Ubound
			End If
			Next rw
			driver.SwitchToDefaultContent
			driver.SwitchToFrame (0)
			If clms = "Accepted" Then
				i = driver.FindElementsByTag("tr").Count
				final = ""
				X = 1
				For Each tr In driver.FindElementsByTag("tr")
					If tr.id = "reportListItemID" Then
'If i = 21 Or i = 22 Or i = 23 Or i = 27 Or i = 28 Or i = 29 Or i = 30 Or i = 34 Or i = 35 Or i = 36 Then
						If i = X Then
							Debug.Print tr.Text
							final = tr.Text & vbCrLf
						End If
'i = i + 1
						X = X + 1
					End If
				Next
			Else
				i = driver.FindElementsByTag("tr").Count
				final = ""
				X = 1
				For Each tr In driver.FindElementsByTag("tr")
					If tr.id = "reportListItemID" Then
'If i = 21 Or i = 22 Or i = 23 Or i = 27 Or i = 28 Or i = 29 Or i = 30 Or i = 34 Or i = 35 Or i = 36 Then
						If i = X Then
							Debug.Print tr.Text
'final = tr.Text & vbCrLf
							final = "Claim not found"
						End If
'i = i + 1
						X = X + 1
					End If
				Next
			End If
'If final = "" Then final = "Claim not found"
			rs!Message = final
			rs.Update
			driver.Window.Close
			driver.SwitchToPreviousWindow
			rs.MoveNext
			
		Wend
		rs.Close
		MsgBox ("Patient Details Done")
		
	End Sub
	
	Private Sub cmdPaymentSearch_Click()
		
		Dim str As String
		Dim rs As New ADODB.Recordset
'Dim sRxnbr As String
		
		rs.Open "Select * from tbl_InputData where [General Recipient Nbr] is not null and paymentsearchdate is null", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
'rs.Close
		
		While Not rs.EOF
			
			driver.Get "https://cda.changehealthcare.com/ERANEW/era/searchPayment.do?isResearch=0&redirect=true&hasHeader=true"
			driver.Window.Maximize
			Sleep 2000
			
			driver.SwitchToDefaultContent
			Dim dd1 As Selenium.SelectElement
			Dim op As Selenium.WebElement
			
			Set dd1 = driver.FindElementByName("period").AsSelect
			dd1.SelectByText "-- ANY --"
'dd1.SelectByText dd1.Options.First.Text
'driver.FindElementByName("period").Count
'driver.FindElementByXPath("//*[@id='basic-search-container']/div/div[6]/select/option[1]").AsSelect.SelectedOption
			
			driver.FindElementsByXPath("//*[@id='searchRadio']")(2).ClickDouble
			Sleep 1000
			driver.FindElementByName("insuredId").SendKeys (rs![General Recipient Nbr])
			driver.FindElementById("paySearch").Click
			Sleep 2000
			
			
			Runext:
			If driver.FindElementsByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr").Count >= 1 Then
				
				If driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr/td").Text <> "Your search returns no results." Then
					For i = 1 To driver.FindElementsByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr").Count
						
'                Debug.Print driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[2]").Text
'                Debug.Print rs![rx nbr]
'                Debug.Print driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[5]").Text
'                Debug.Print Format(rs![Fill Entered Dttm], "mm/dd/yy")
'(driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[2]").Text Like "*" & rs![rx nbr] & "* ")
						If (driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[2]").Text Like "" & rs![rx nbr] & "*") And (driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[5]").Text = Format(rs![Fill Entered Dttm], "mm/dd/yy")) Then 'And
							
'rs!Icn = driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[4]").Text
'rs!Providername = driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[6]").Text
							rs!BatchID = driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[8]").Text
							
							On Error Resume Next
							If Err.Number <> 7 Then
								driver.FindElementByXPath("/html/body/div[1]/div[1]/div/form/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[2]/a").Click
								Sleep 5000
								driver.FindElementByXPath("/html/body/div[1]/div[1]/div/form/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[2]/a").Click
								Sleep 3000
' driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[2]").Click
'Sleep 2000
								
							End If
							Sleep 5000
							rs!Providername = driver.FindElementByXPath("//*[@id='content-holder']/table[1]/tbody/tr[2]/td[1]/table/tbody/tr[1]/td").Text
							rs!Payment = Replace(driver.FindElementByXPath("/html/body/div/table[1]/tbody/tr[2]/td[3]/table/tbody/tr[2]/td").Text, "Payment #: ", "")
							rs![Payment Date] = Replace(driver.FindElementByXPath("//*[@id='content-holder']/table[1]/tbody/tr[2]/td[3]/table/tbody/tr[3]/td").Text, "Payment Date:", "")
							rs![Payment Amount] = Replace(driver.FindElementByXPath("/html/body/div/table[1]/tbody/tr[2]/td[3]/table/tbody/tr[4]/td").Text, "Payment Amount:", "")
							rs![Total PLB Adjustments] = Replace(driver.FindElementByXPath("//*[@id='content-holder']/table[1]/tbody/tr[2]/td[3]/table/tbody/tr[5]/td").Text, "Total PLB Adjustments:", "")
							rs![CCN/ICN] = driver.FindElementByXPath("//*[@id='content-holder']/table[2]/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[6]/td[2]").Text
							rs![Total Charge] = driver.FindElementByXPath("//*[@id='content-holder']/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[1]/td[2]").Text
							rs![Total Payment] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[2]/td[2]").Text
							rs![Total Contractual] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[3]/td[2]").Text
							rs![Total Deductible] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[4]/td[2]").Text
							rs![Total Co-insurance] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[5]/td[2]").Text
							rs![Total Co-Payment] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[5]/td[2]").Text
							rs![Service Date] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[5]/td[2]").Text
							rs![Network ID] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[7]/td[2]").Text
							driver.FindElementByXPath("/html/body/div/div[9]/img").Click
							Sleep 2000
							rs![Proc Code1] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[3]").Text
							rs![Proc Code2] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[4]/td[3]").Text 'driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[3]").Text
							If Len(rs![Proc Code2]) < 2 Then rs![Proc Code2] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[3]/td[3]").Text
							
							rs!Charges1 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[9]").Text
							rs!Charges2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[4]/td[9]").Text
							If Len(rs!Charges2) < 2 Then rs!Charges2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[3]/td[9]").Text
							
							rs!Allowed1 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[10]").Text
							rs!Allowed2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[4]/td[10]").Text
							If Len(rs!Allowed2) < 2 Then rs!Allowed2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[3]/td[10]").Text
							
							rs!Payment1 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[11]").Text
							rs!Payment2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[4]/td[11]").Text
							If Len(rs!Payment2) < 2 Then rs!Payment2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[3]/td[11]").Text
							
							rs![Adj Amt1] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[12]").Text
							rs![Adj Amt2] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[4]/td[12]").Text
'If Len(rs!Payment2) < 2 Then rs!Payment2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[3]/td[11]").Text
							
							rs![Adj Codes/Descriptions1] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[13]").Text
							rs![Adj Codes/Descriptions2] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[13]").Text
							rs![Adj Codes/Descriptions3] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[14]").Text
							rs![Adj Codes/Descriptions4] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[14]").Text
							
							driver.WaitForScript ("return document.readyState")
							
							GoTo RunextRs
						End If
						Next i
					End If
				End If
				
				On Error Resume Next
				If driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[2]/div[4]/a[2]").Text = "Next >>" Then
					If Err.Number <> 7 Then
						driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[2]/div[4]/a[2]").Click
						Sleep 2000
						On Error GoTo 0
						GoTo Runext
					End If
				End If
				On Error GoTo 0
				RunextRs:
				rs!PaymentSearchDate = Date
				rs.Update
				rs.MoveNext
			Wend
			rs.Close
			MsgBox ("Pocess Done")
			
''    Dim str As String
''    Dim rs As New ADODB.Recordset
''    Dim sRxnbr As String
''
''        rs.Open "Select * from tbl_InputData", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
''        'rs.Close
''
''While Not rs.EOF
''
''        driver.Get "https://cda.changehealthcare.com/ERANEW/era/searchPayment.do?isResearch=0&redirect=true&hasHeader=true"
''        driver.Window.Maximize
''        Sleep 2000
''
''        driver.SwitchToDefaultContent
''        Dim dd1 As Selenium.SelectElement
''        Dim op As Selenium.WebElement
''
''        Set dd1 = driver.FindElementByName("period").AsSelect
''        dd1.SelectByText "-- ANY --"
''       'dd1.SelectByText dd1.Options.First.Text
''       'driver.FindElementByName("period").Count
''       'driver.FindElementByXPath("//*[@id='basic-search-container']/div/div[6]/select/option[1]").AsSelect.SelectedOption
''
''        driver.FindElementsByXPath("//*[@id='searchRadio']")(2).ClickDouble
''        Sleep 1000
''        driver.FindElementByName("insuredId").SendKeys (rs![General Recipient Nbr])
''        driver.FindElementById("paySearch").Click
''        Sleep 2000
''
''
''Runext:
''        If driver.FindElementsByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr").Count > 1 Then
''
''            For i = 1 To driver.FindElementsByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr").Count
''
''               If driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[2]").Text = sRxnbr Then
''
''                    'rs!Icn = driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[4]").Text
''                    'rs!Providername = driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[6]").Text
''                    rs!BatchID = driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[8]").Text
''
''                    On Error Resume Next
''                    If Err.Number <> 7 Then
''                        driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[1]/div/table/tbody/tr[" & i & "]/td[2]").Click
''                    End If
''                    Sleep 2000
''                    rs!Providername = driver.FindElementByXPath("//*[@id='content-holder']/table[1]/tbody/tr[2]/td[1]/table/tbody/tr[1]/td").Text
''                    rs!Payment = Replace(driver.FindElementByXPath("/html/body/div/table[1]/tbody/tr[2]/td[3]/table/tbody/tr[2]/td").Text, "Payment #: ", "")
''                    rs![Payment Date] = Replace(driver.FindElementByXPath("//*[@id='content-holder']/table[1]/tbody/tr[2]/td[3]/table/tbody/tr[3]/td").Text, "Payment Date:", "")
''                    rs![Payment Amount] = Replace(driver.FindElementByXPath("/html/body/div/table[1]/tbody/tr[2]/td[3]/table/tbody/tr[4]/td").Text, "Payment Amount:", "")
''                    rs![Total PLB Adjustments] = Replace(driver.FindElementByXPath("//*[@id='content-holder']/table[1]/tbody/tr[2]/td[3]/table/tbody/tr[5]/td").Text, "Total PLB Adjustments:", "")
''                    rs![CCN/ICN] = driver.FindElementByXPath("//*[@id='content-holder']/table[2]/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[6]/td[2]").Text
''                    rs![Total Charge] = driver.FindElementByXPath("//*[@id='content-holder']/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[1]/td[2]").Text
''                    rs![Total Payment] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[2]/td[2]").Text
''                    rs![Total Contractual] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[3]/td[2]").Text
''                    rs![Total Deductible] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[4]/td[2]").Text
''                    rs![Total Co-insurance] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[5]/td[2]").Text
''                    rs![Total Co-Payment] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[3]/table/tbody/tr[5]/td[2]").Text
''                    rs![Service Date] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[5]/td[2]").Text
''                    rs![Network ID] = driver.FindElementByXPath("/html/body/div/table[2]/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[7]/td[2]").Text
''                    driver.FindElementByXPath("/html/body/div/div[9]/img").Click
''                    Sleep 2000
''                    rs![Proc Code1] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[3]").Text
''                    rs![Proc Code2] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[3]").Text
''                    rs!Charges1 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[9]").Text
''                    rs!Charges2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[9]").Text
''                    rs!Allowed1 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[10]").Text
''                    rs!Allowed2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[10]").Text
''                    rs!Payment1 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[11]").Text
''                    rs!Payment2 = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[11]").Text
''                    rs![Adj Amt1] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[12]").Text
''                    rs![Adj Amt2] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[12]").Text
''                    rs![Adj Codes/Descriptions1] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[13]").Text
''                    rs![Adj Codes/Descriptions2] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[13]").Text
''                    rs![Adj Codes/Descriptions3] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[1]/td[14]").Text
''                    rs![Adj Codes/Descriptions4] = driver.FindElementByXPath("/html/body/div/div[10]/div/table/tbody/tr[2]/td[14]").Text
''
''                    driver.WaitForScript ("return document.readyState")
''
''                    GoTo RunextRs
''               End If
''            Next i
''        End If
''
''        On Error Resume Next
''        If driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[2]/div[4]/a[2]").Text = "Next >>" Then
''            If Err.Number <> 7 Then
''                driver.FindElementByXPath("//*[@id='searchForm']/div[3]/div[2]/div[4]/a[2]").Click
''                Sleep 2000
''                On Error GoTo 0
''                GoTo Runext
''            End If
''       End If
''       On Error GoTo 0
''RunextRs:
''        rs.Update
''        rs.MoveNext
''  Wend
''        rs.Close
''        MsgBox ("Pocess Done")
		End Sub
		
		
'Export Output Report
		Private Sub cmdReport_Click()
			DoCmd.RunMacro ("McrExport")
			MsgBox ("Report Exported in Current Path!")
		End Sub
		
		Private Sub cmdRx_nbr_Click()
			
			Dim str As String
			Dim clms As String
			Dim rs As New ADODB.Recordset
			
'DoCmd.OpenQuery ("Qry_Rxnbr_addZero")
			
'     r = "No Records found based on your search criteria. Please modify the search criteria and re-submit."
			
'   rs.Open "Select * from tbl_InputData where Message = '" & r & "'", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
'   'rs.Open "Select * from tbl_InputData where Message = """ & r & """"", CurrentProject.Connection, adOpenDynamic, adLockOptimistic"
''   rs.Open "Select * from tbl_InputData where Message is not null", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
'
'    rs.Open "Select * from tbl_InputData where Message is not null, CurrentProject.Connection, adOpenDynamic, adLockOptimistic"
'    rs.Open "Select * from tbl_InputData where Message = 'No Records found based on your search criteria. Please modify the search criteria and re-submit.'", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
			rs.Open "Select * from tbl_InputData where Message like '%No Records%' and RXSearchDate is null", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
			While Not rs.EOF
				
				StartDate = CDate(rs![Fill Entered Dttm]) - 30
				EndDate = CDate(rs![Fill Entered Dttm]) + 30
				sdate = StartDate
				edate = EndDate
				sFdt = Format(StartDate, "Mmm") & " " & Format(StartDate, "d") & ", " & Format(StartDate, "YYYY")
				sTdt = Format(EndDate, "Mmm") & " " & Format(EndDate, "d") & ", " & Format(EndDate, "YYYY")
				
				If Len(rs![rx nbr]) = 7 Then
					sRxnbr = rs![rx nbr]
				ElseIf Len(rs![rx nbr]) = 6 Then
					sRxnbr = "0" & rs![rx nbr]
				ElseIf Len(rs![rx nbr]) = 5 Then
					sRxnbr = "00" & rs![rx nbr]
				ElseIf Len(rs![rx nbr]) = 4 Then
					sRxnbr = "000" & rs![rx nbr]
				End If
'        str = "https://access.emdeon.com/ProviderVision/provider/getJspxReport.jspx?id=PR_1001_Insured&event_name_id="
'        str = str & "Patient Search&from=" & sFdt & "&to=" & sTdt & "&npi=&date.type=SERVICE_FROM&insured.id=" & rs![General Recipient Nbr]
'        str = str & "&dateFrom=" & StartDate & "&dateTo=" & EndDate & "&is_from_search_page=yes&fromPage=search&searchType="
'        str = str & "Patient Search&parentURL=/ProviderVision/provider/getIndexPage.jspx&random=1651002180232"
'
'https://access.emdeon.com/ProviderVision/provider/getJspxReport.jspx?id=PR_1001_Insured&event_name_id=
'Patient%20Search&from=Jun%204,%202021&to=Jul%204,%202022&npi=&date.type=SERVICE_FROM&pcn.id=2985798
'&dateFrom=06/04/2021&dateTo=07/04/2022&is_from_search_page=yes&fromPage=search&searchType=
'Patient%20Search&parentURL=/ProviderVision/provider/getIndexPage.jspx&random=1657024421903
				
				str = "https://access.emdeon.com/ProviderVision/provider/getJspxReport.jspx?id=PR_1001_Insured&event_name_id="
				str = str & "Patient Search&from=" & sFdt & "&to=" & sTdt & "&npi=&date.type=SERVICE_FROM&pcn.id=" & sRxnbr
				str = str & "&dateFrom=" & StartDate & "&dateTo=" & EndDate & "&is_from_search_page=yes&fromPage=search&searchType="
				str = str & "Patient Search&parentURL=/ProviderVision/provider/getIndexPage.jspx&random=1657024421903"
				
				driver.ExecuteScript "window.open(arguments[0])", str
				Sleep (2000)
				driver.Window.Maximize
				driver.SwitchToNextWindow
				driver.SwitchToDefaultContent
				
				Dim Found As Integer
				Found = 0
				i = 1
				For Each rw In driver.FindElementById("reportContentTable").FindElementsByTag("tr")
					Debug.Print rw.Text
					On Error Resume Next
					cln = Split(rw.Text, " ")
					If Err.Number <> 438 Then
						If UBound(cln) > 3 Then
							If InStr(1, cln(3), "-") > 0 Then
								If CLng(Trim(Left(cln(3), InStr(1, cln(3), "-") - 1))) Like " " & Trim(sRxnbr) & "*" And Trim(cln(4)) = Format(rs!WebSDL_SOLD, "mm/dd/yyyy") And Trim(cln(4)) = Format(rs![Fill Entered Dttm], "mm/dd/yyyy") Then
									clms = driver.FindElementByXPath("//*[@id=""reportListItemID" & i & """]/td[6]").Text
									driver.FindElementByXPath("//*[@id=""reportListItemID" & i & """]/td[8]/a").Click
									Sleep 2000
									Found = 1
									Exit For
								End If
							Else
								If CLng(Trim(cln(3))) = Trim(sRxnbr) And Trim(cln(4)) = Format(rs!WebSDL_SOLD, "mm/dd/yyyy") Then
									clms = driver.FindElementByXPath("//*[@id=""reportListItemID" & i & """]/td[6]").Text
									driver.FindElementByXPath("//*[@id=""reportListItemID" & i & """]/td[8]/a").Click
									Sleep 2000
									Found = 1
									Exit For
								End If
							End If 'instr
							i = i + 1
						End If 'Ubound
					End If
					Next rw
					driver.SwitchToDefaultContent
					driver.SwitchToFrame (0)
					If clms = "Accepted" Then
						i = driver.FindElementsByTag("tr").Count
						final = ""
						X = 1
						For Each tr In driver.FindElementsByTag("tr")
							If tr.id = "reportListItemID" Then
'If i = 21 Or i = 22 Or i = 23 Or i = 27 Or i = 28 Or i = 29 Or i = 30 Or i = 34 Or i = 35 Or i = 36 Then
								If i = X Then
									Debug.Print tr.Text
									final = tr.Text & vbCrLf
								End If
'i = i + 1
								X = X + 1
							End If
						Next
					Else
						i = driver.FindElementsByTag("tr").Count
						final = ""
						X = 1
						For Each tr In driver.FindElementsByTag("tr")
							If tr.id = "reportListItemID" Then
'If i = 21 Or i = 22 Or i = 23 Or i = 27 Or i = 28 Or i = 29 Or i = 30 Or i = 34 Or i = 35 Or i = 36 Then
								If i = X Then
									Debug.Print tr.Text
									final = tr.Text & vbCrLf
								End If
'i = i + 1
								X = X + 1
							End If
						Next
					End If
					rs!Message = final
					rs!RXSearchDate = Date
					rs.Update
					driver.Window.Close
					driver.SwitchToPreviousWindow
					rs.MoveNext
					
				Wend
				rs.Close
				MsgBox ("Rx numbers searching completed")
				
				
			End Sub
			
			
			
