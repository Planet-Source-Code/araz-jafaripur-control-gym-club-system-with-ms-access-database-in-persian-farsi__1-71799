Attribute VB_Name = "DB_Func"
Dim Connect As New ADODB.Connection
Dim RS As New ADODB.Recordset
Public Sub ConnectToDb(adoObj As Adodc, AdoRec As String, IsSql As Boolean)
'    On Error Resume Next
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\Data.dat;Persist Security Info=False; Jet OLEDB:Database Password = 123456Jafaripur"
    If IsSql Then
        adoObj.CommandType = adCmdText
    Else
        adoObj.CommandType = adCmdTable
    End If
    adoObj.RecordSource = AdoRec
    adoObj.Refresh
End Sub
Private Sub ConDB()
    On Error Resume Next
    RS.Close
    Connect.Close
    Connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\Data.dat;Persist Security Info=False; Jet OLEDB:Database Password = 123456Jafaripur"
End Sub
Public Sub DatagridColumnAutoResize(ByRef oDataGrid As DataGrid, _
                                    ByRef oForm As Form)
Dim I As Integer, iMax As Integer
Dim t As Integer, tMax As Integer
Dim iWidth As Integer
Dim vBMark As Variant
Dim aWidth As Variant
Dim cText As String
Dim oFont As Font
    On Error Resume Next
    oFont = oForm.Font
    oForm.Font = oDataGrid.Font
    iMax = oDataGrid.Columns.Count - 1
    ReDim aWidth(iMax)
    For I = 0 To iMax
        aWidth(I) = 0
        DoEvents
    Next
    tMax = oDataGrid.VisibleRows - 1
    For t = 0 To tMax
        vBMark = oDataGrid.GetBookmark(t)
        For I = 0 To iMax
            cText = oDataGrid.Columns(I).CellText(vBMark)
            iWidth = oForm.TextWidth(cText)
            If iWidth + ((12 * Len(cText)) + 220) > aWidth(I) Then
                aWidth(I) = iWidth + ((12 * Len(cText)) + 220)
            End If
            If t = 0 Then
                iWidth = oForm.TextWidth(oDataGrid.Columns(I).Caption)
                If iWidth + ((12 * Len(cText)) + 220) > aWidth(I) Then
                    aWidth(I) = iWidth + ((12 * Len(cText)) + 220)
                End If
            End If
            DoEvents
        Next
        DoEvents
    Next
    For I = 0 To iMax
        oDataGrid.Columns(I).Width = aWidth(I)
        DoEvents
    Next
    oForm.Font = oFont
End Sub
Public Sub AdminToCbo(CboAdmin As ComboBox)
Call ConDB
RS.Open "SELECT UserName FROM Admin", Connect, adOpenStatic, adLockOptimistic
While Not RS.EOF
    CboAdmin.AddItem RS.Fields("UserName")
    RS.MoveNext
    DoEvents
Wend
RS.Close
Connect.Close
Set RS = Nothing
Set Connect = Nothing
End Sub
Public Function GetPay(ID As Long) As Integer
Call ConDB
RS.Open "SELECT ID,UID,DateGive_S FROM Pay WHERE UID = " & ID & " ORDER BY ID ASC", Connect, adOpenStatic, adLockOptimistic
If RS.EOF And RS.BOF Then
    GetPay = 1
Else
    Dim DateGive As String
    Dim TodayDate As String
    Dim Result As Integer
    RS.MoveLast
    DateGive = RS.Fields("DateGive_S")
    DateGive = Shamsi.IncreaseDate_Custom(DateGive, 30)
    TodayDate = Shamsi.Convert_Date(Date, Gregorian_, HijriShamsi_)
    Call Shamsi.DateCompare(DateGive, TodayDate, Result)
    GetPay = Result
End If
RS.Close
Connect.Close
Set RS = Nothing
Set Connect = Nothing
End Function
Public Function EnterCode(ID As Long, Code As Integer) As Integer
Call ConDB
RS.Open "SELECT UID,GiveNum FROM UserLogin WHERE GiveNum = 0 AND UID = " & ID, Connect, adOpenStatic, adLockOptimistic
If RS.EOF And RS.BOF Then
    EnterCode = 0
Else
    EnterCode = -1
    RS.Close
    Connect.Close
    Set RS = Nothing
    Set Connect = Nothing
    Exit Function
End If
RS.Close
RS.Open "SELECT KNum,GiveNum FROM UserLogin WHERE GiveNum = 0 AND KNum = " & Code, Connect, adOpenStatic, adLockOptimistic
If RS.EOF And RS.BOF Then
    EnterCode = 0
Else
    EnterCode = -2
End If
RS.Close
Connect.Close
Set RS = Nothing
Set Connect = Nothing
End Function
