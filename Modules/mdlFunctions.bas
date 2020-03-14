Attribute VB_Name = "mdlFunctions"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

Private Const LVM_SETCOLUMNWIDTH = &H101E
Private Const WM_SETREDRAW As Long = &HB&

Public Sub ClearVismVars()
    ' Clear data in VisM Vars. This avoids cache problems
    With mdiMain.VisM
        .P0 = Empty: .P1 = Empty: .P2 = Empty: .P3 = Empty: .P4 = Empty
        .P5 = Empty: .P6 = Empty: .P7 = Empty: .P8 = Empty: .P9 = Empty
    End With
End Sub

Public Function IndexToMenu(iIndex As Integer) As Object

    ' It is necessary to associate a menu with an index.
    ' Otherwise, we cannot dynamically add options to the menus.
    With mdiMain
        Select Case iIndex
            Case 1: Set IndexToMenu = .mnSubMain1
            Case 2: Set IndexToMenu = .mnSubMain2
        End Select
    End With
End Function

Public Sub AddSubMenu(sCaption As String, oMenu As Object, Optional sTag As String = Empty)
    Dim iIndex As Integer
    
    iIndex = oMenu.Count
    If iIndex = 1 And oMenu(iIndex).Caption = Empty Then
        iIndex = iIndex
    Else
        Load oMenu(iIndex + 1)
        iIndex = iIndex + 1
    End If
    
    oMenu(iIndex).Caption = sCaption
    oMenu(iIndex).Tag = sTag
    oMenu(iIndex).Visible = True
End Sub

Public Function GetWidth(sChar As String) As Long

    ' I think it's easier to remember three or more letters instead of numbers.
    Select Case UCase(sChar)
        Case "S": GetWidth = 1335
        Case "M": GetWidth = 2175
        Case "L": GetWidth = 3015
    End Select
End Function

Public Function FileExists(sFullPath As String) As Boolean
    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(sFullPath)
End Function

Function Export2Excel( _
    sFileName As String, _
    lsw As ListView, _
    Optional ProgressBar As ProgressBar, _
    Optional SheetIndex As Integer = 1) As Boolean
      
    On Error GoTo error_Handler
      
    Dim obj_Excel       As Object ' For Excel.Application
    Dim obj_Libro       As Object
      
    Dim iCol            As Integer
    Dim iRow            As Long
    
    Set obj_Excel = CreateObject("Excel.Application")
    Set obj_Libro = obj_Excel.Workbooks.Add
  
    With obj_Libro

        If Not ProgressBar Is Nothing Then
            ProgressBar.Max = lsw.ListItems.Count
            If Not ProgressBar.Visible Then ProgressBar.Visible = True
        End If

        With .Sheets(SheetIndex)
        
            .Cells.Font.Size = 8
            .Cells.Font.Name = "Courier New"

            iCol = 1
            iRow = 1
            For iCol = 1 To lsw.ColumnHeaders.Count
                .Cells(iRow, iCol).Font.Bold = True
                .Cells(iRow, iCol) = lsw.ColumnHeaders(iCol).Text
            Next

            For iRow = 1 To lsw.ListItems.Count
                iCol = 1

                If lsw.ColumnHeaders(iCol).Tag = "NUMBER" Then
                    
                    .Cells(iRow + 1, iCol).NumberFormat = "#,##0.00"
                    .Cells(iRow + 1, iCol) = Replace(Replace(lsw.ListItems(iRow), ".", ""), ",", ".")
                ElseIf lsw.ColumnHeaders(iCol).Tag = "DATE" Then
                    
                    .Cells(iRow + 1, iCol).Value = CDate(lsw.ListItems(iRow))
                Else
                    
                    .Cells(iRow + 1, iCol).NumberFormat = "General"
                    .Cells(iRow + 1, iCol) = lsw.ListItems(iRow)
                End If

                For iCol = 1 To lsw.ColumnHeaders.Count - 1

                    If Trim(lsw.ListItems(iRow).SubItems(iCol)) <> "" Then
                        If lsw.ColumnHeaders(iCol + 1).Tag = "NUMBER" Then
                            
                            .Cells(iRow + 1, iCol + 1).NumberFormat = "#,##0.00"
                            .Cells(iRow + 1, iCol + 1) = Replace(Replace(lsw.ListItems(iRow).SubItems(iCol), ".", ""), ",", ".")
                        ElseIf lsw.ColumnHeaders(iCol + 1).Tag = "DATE" Then
                            
                            .Cells(iRow + 1, iCol + 1).Value = CDate(lsw.ListItems(iRow).SubItems(iCol))
                        Else
                            
                            .Cells(iRow + 1, iCol + 1).NumberFormat = "General"
                            .Cells(iRow + 1, iCol + 1) = lsw.ListItems(iRow).SubItems(iCol)
                        End If
                    Else
                        
                        .Cells(iRow + 1, iCol + 1).NumberFormat = "General"
                        .Cells(iRow + 1, iCol + 1) = lsw.ListItems(iRow).SubItems(iCol)
                    End If
                 Next
  
                 If Not ProgressBar Is Nothing Then
                     ProgressBar.Value = ProgressBar.Value + 1
                 End If
            Next
            
            .Columns("A:F").EntireColumn.AutoFit
            obj_Excel.ActiveWindow.SplitRow = 1
            obj_Excel.ActiveWindow.FreezePanes = True
        End With
    End With
    
    obj_Excel.Visible = True
    
    ' Variables cleaning
    Set obj_Libro = Nothing
    Set obj_Excel = Nothing
    
    ' OK!
    Export2Excel = True
      
    If Not ProgressBar Is Nothing Then
       ProgressBar.Value = 0
       ProgressBar.Visible = False
    End If
    
    Exit Function
error_Handler:
      
    Export2Excel = False
    MsgBox Err.Description, vbCritical, vbOKOnly, App.ProductName
      
    On Error Resume Next
    Set obj_Libro = Nothing
    Set obj_Excel = Nothing
    
    If Not ProgressBar Is Nothing Then
        ProgressBar.Value = 0
    End If
End Function

Public Function Export2Calc(lsw As ListView, Optional pbBarra As ProgressBar = Nothing) As Boolean
    Dim ServiceManager As Object, Desktop As Object, Document As Object, Feuille As Object
    Dim Plage As Object, objCols As Object, objCol As Object, objRows As Object, objRow As Object
    Dim oCell As Object
    Dim PrintArea(0)
    Dim PrintArgs(2)
    Dim i As Long, j As Long
    Dim sTemp As String
    
    Const TVA = "1.196"
    
    On Error GoTo error
    
    ' Calc instance
    Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
    Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
    
    ' Requiered configuration
    Dim args(1) As Object
    Set args(0) = ServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    args(0).Name = "Hidden"
    args(0).Value = True
    
    ' Create file
    Set Document = Desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, args)
    Set Feuille = Document.getSheets().getByIndex(0)

    With lsw
    
        ' ProgressBar configuration
        If Not pbBarra Is Nothing Then
            pbBarra.Max = lsw.ListItems.Count + 1
            If Not pbBarra.Visible Then pbBarra.Visible = True
            pbBarra.Min = 0
            pbBarra.Value = 0
        End If
        
        ' Header
        For i = 1 To .ColumnHeaders.Count
            Feuille.getCellByPosition(i - 1, 0).CharWeight = 150
            Call Feuille.getCellByPosition(i - 1, 0).setPropertyValue("CharFontName", "Courier New")
            Call Feuille.getCellByPosition(i - 1, 0).setPropertyValue("CharHeight", 10)
            Call Feuille.getCellByPosition(i - 1, 0).setFormula(.ColumnHeaders(i).Text)
        Next i

        ' Data
        For i = 1 To .ListItems.Count
            For j = 1 To .ColumnHeaders.Count
                
                sTemp = ""
                If j = 1 Then
                    sTemp = .ListItems(i).Text
                    If IsNumeric(sTemp) Then sTemp = Replace(sTemp, ",", ".")
                Else
                    sTemp = .ListItems(i).SubItems(j - 1)
                    If IsNumeric(sTemp) Then sTemp = Replace(sTemp, ",", ".")
                End If

                Call Feuille.getCellByPosition(j - 1, i).setPropertyValue("CharFontName", "Courier New")
                Call Feuille.getCellByPosition(j - 1, i).setPropertyValue("CharHeight", 10)
                
                Set oCell = Feuille.getCellByPosition(j - 1, i)
                oCell.String = sTemp
                Set oCell = Nothing
            Next j
            If Not pbBarra Is Nothing Then
                pbBarra.Value = pbBarra.Value + 1
            End If
        Next i
    End With
    
    Set objCols = Feuille.getColumns
    
    For i = 1 To lsw.ColumnHeaders.Count
        objCols.getByIndex(i - 1).optimalWidth = True
    Next i
    
    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
    Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
    Call Document.currentController.freezeAtPosition(0, 1)

    ' Variables cleaning
    Set Plage = Nothing
    Set Feuille = Nothing
    Set Document = Nothing
    Set Desktop = Nothing
    Set ServiceManager = Nothing
    
    If Not pbBarra Is Nothing Then
        If pbBarra.Visible Then pbBarra.Visible = False
    End If
    
    Export2Calc = True
    
    Exit Function
error:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
    Export2Calc = False
    
    ' Variables cleaning
    If Not Plage Is Nothing Then Set Plage = Nothing
    If Not Feuille Is Nothing Then Set Feuille = Nothing
    If Not Document Is Nothing Then Set Document = Nothing
    If Not Desktop Is Nothing Then Set Desktop = Nothing
    If Not ServiceManager Is Nothing Then Set ServiceManager = Nothing
End Function

Function Separate(sText As String, sSeparator As String, iPosition As Integer) As String
    ' Returns a string that is between a specific character.
    ' Sometimes it is simpler than converting a string to an array
    
    Dim II As Integer
    Dim Pos1 As Integer
    Dim Pos2 As Integer
    
    Pos1 = 1
    Pos2 = Len(sText)
    For II = 1 To iPosition
        If II <> 1 Then Pos1 = Pos2 + Len(sSeparator) + 1
        Pos2 = InStr(Pos1, sText, sSeparator) - 1
        If Pos2 < 0 Then Pos2 = Len(sText)
    Next
    If Pos2 < 0 Then Pos2 = Len(sText)
    If Pos2 >= Pos1 Then
        Separate = Mid(sText, Pos1, Pos2 - Pos1 + 1)
    Else
        Separate = ""
    End If
End Function

Public Sub AutosizeColumns(ByVal TargetListView As ListView)
 
    ' Fit ListView columns to the width of the longest row content
    Const SET_COLUMN_WIDTH  As Long = 4126
    Const AUTOSIZE_USEHEADER As Long = -2
     
    Dim lngColumn As Long
     
    For lngColumn = 0 To (TargetListView.ColumnHeaders.Count - 1)
     
        Call SendMessage(TargetListView.hwnd, _
        SET_COLUMN_WIDTH, _
        lngColumn, _
        ByVal AUTOSIZE_USEHEADER)
     
    Next lngColumn
 
End Sub

Public Sub OrderListView(ByVal Handler As Form, ByVal lsw As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim i As Long
    Dim sFormat As String
    Dim strData() As String
    Dim lColumn As Long
      
    With lsw
        Call SendMessage(Handler.hwnd, WM_SETREDRAW, 0&, 0&)
        lColumn = ColumnHeader.Index - 1
    
        Select Case UCase$(ColumnHeader.Tag) ' Format Type
            Case "DATE"
                sFormat = "YYYYMMDDHhNnSs"
            With .ListItems
                If (lColumn > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(lColumn)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    sFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    sFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
              
            With .ListItems
                If (lColumn > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(lColumn)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                End If
            End With

        Case "NUMBER"

            sFormat = String(30, "0") & "." & String(30, "0")
                  
            With .ListItems
                If (lColumn > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(lColumn)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        sFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        sFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        sFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        sFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With

            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
              
            With .ListItems
                If (lColumn > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(lColumn)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                End If
            End With
          
        Case Else
                      
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
              
        End Select
      
    End With
      
    Call SendMessage(Handler.hwnd, WM_SETREDRAW, 1&, 0&)
    lsw.Refresh
      
End Sub

Public Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function

Public Sub PaintObj(obj As TextBox, Optional lColor As Long = &HBBFFFF, Optional bSelect As Boolean = True)

    If bSelect Then
        obj.SelStart = 0
        obj.SelLength = Len(obj.Text)
    End If
    
    obj.BackColor = lColor
End Sub

Public Sub EnterTab(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        SendKeys "+{TAB}"
    End If
End Sub

Sub ForceCaps(ByRef KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Sub AutoGenerateForm(oMnu As Object, iIndex As Integer)
    If oMnu(iIndex).Tag = Empty Then Exit Sub
    
    Dim frm As New frmGenericReport
    Set frm = New frmGenericReport
    With mdiMain.VisM
        .P0 = oMnu(iIndex).Tag
        .Execute ("D GETFORM^CADR(P0,.P1,.P2,.P3)")

        If .P1 = Empty Or .P2 = Empty Or .P3 = Empty Then
            MsgBox "There was an error opening the requested form.", vbCritical + vbOKOnly, "Dynamic Report"
            Exit Sub
        End If

        frm.Caption = .P1
        frm.sQuery = .P2
        frm.sGlobal = Separate(.P2, "^", 2)
        frm.sComp = .P3
        frm.sIdForm = oMnu(iIndex).Tag

        frm.Show
    End With
    Set frm = Nothing
End Sub

Public Function Export2CSV(lsw As ListView, Optional sPath As String = "", Optional bHeader As Boolean = True) As Boolean
    Dim i As Long, j As Long
    Dim iFile As Integer
    Dim sLine As String
    
    If sPath = "" Then sPath = App.Path & "/Report_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    
    iFile = FreeFile
    Open sPath For Output As #iFile
    
    sLine = Empty
    If bHeader Then
        For i = 1 To lsw.ColumnHeaders.Count
            If sLine <> "" Then sLine = sLine & ";"
            sLine = sLine & lsw.ColumnHeaders(i).Text
        Next i
        Print #iFile, sLine
    End If
    
    For i = 1 To lsw.ListItems.Count
        sLine = Empty
    
        For j = 1 To lsw.ColumnHeaders.Count
            If sLine <> "" Then sLine = sLine & ";"
            
            If j = 1 Then
                sLine = sLine & lsw.ListItems(i).Text
            Else
                sLine = sLine & lsw.ListItems(i).SubItems(j - 1)
            End If

        Next j
        
        Print #iFile, sLine
    Next i
    
    Close #iFile
    
    If FileExists(sPath) Then
        OpenFile sPath, False
    End If
    
    Export2CSV = True
    Exit Function
error:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Export2CSV"
    Export2CSV = False
End Function

Public Sub OpenFile(strPath As String, Optional bShowError As Boolean = True)
    Dim lRet As Long
    lRet = ShellExecute(vbNull, "", strPath, "", "", 1)

    If bShowError And lRet <= 32 Then
        MsgBox "Could not open the file", vbCritical + vbOKOnly, "Error!"
    End If
End Sub
