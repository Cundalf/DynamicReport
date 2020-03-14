VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGenericReport 
   Caption         =   "Generic Report"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGenericForm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   12135
   Begin VB.Frame frFilters 
      Caption         =   "Filters"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7455
      Begin VB.CheckBox chk 
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   164429825
         CurrentDate     =   42857
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   6000
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame frList 
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   11895
      Begin MSComctlLib.ListView lswList 
         Height          =   6135
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   10821
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton btnExport 
      Caption         =   "Export"
      Height          =   315
      Left            =   10440
      TabIndex        =   0
      Top             =   7560
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "frmGenericReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sIdForm As String
Public sQuery As String
Public sGlobal As String
Public sComp As String

Dim oComponents() As Object

Private Sub btnSearch_Click()
    Dim sValor As String, sVariables As String
    Dim sVar() As String, sColumnas() As String
    Dim i As Integer
    Dim Columna As ColumnHeader
    Dim Item As ListItem
    
    Me.Enabled = False
    
    ' Clear all data
    lswList.ColumnHeaders.Clear
    lswList.ListItems.Clear
    
    ' I assign each component to a VisM variable
    ClearVismVars
    With mdiMain.VisM
        .Execute ("K")
        
        sValor = Empty
        For i = LBound(oComponents) To UBound(oComponents)
            
            If i > LBound(oComponents) Then sValor = sValor & "^"
            Select Case oComponents(i).Tag
                Case "TXT":
                    sValor = sValor & oComponents(i).Text
                Case "CBO":
                    If oComponents(i).ListIndex > -1 Then
                        sValor = sValor & oComponents(i).Text
                    Else
                        sValor = sValor & Empty
                    End If
                Case "DTP":
                    sValor = sValor & Format(oComponents(i).Value, "yyyymmdd")
                Case "CHK":
                    sValor = sValor & oComponents(i).Value
            End Select
        Next i

        If UBound(oComponents) = 0 And sValor = "" Then
            ReDim sVar(0 To 0)
            sVar(0) = ""
        Else
            sVar = Split(sValor, "^")
        End If
        sVariables = Empty
        

        For i = LBound(sVar) To UBound(sVar)
            Select Case i
                Case "0": .P0 = sVar(i)
                Case "1": .P1 = sVar(i)
                Case "2": .P2 = sVar(i)
                Case "3": .P3 = sVar(i)
                Case "4": .P4 = sVar(i)
                Case "5": .P5 = sVar(i)
                Case "6": .P6 = sVar(i)
                Case "7": .P7 = sVar(i)
            End Select
            
            If sVariables <> "" Then sVariables = sVariables & ","
            sVariables = sVariables & "P" & CStr(i)
        Next i
        
        .P8 = USER
        
        ' Execute Query
        ' The last two variables are ALWAYS the USER and a variable that returns the COLUMNS (String)
        
        .Execute ("D " & sQuery & "(" & sVariables & ",P8,.P9)")
        
        ' I always need to have the columns
        If .P9 = "" Then
            MsgBox "Could not get data." & vbCrLf & .ErrorName, vbCritical + vbOKOnly, Me.Caption
            Me.Enabled = True
            Exit Sub
        End If
        
        ' I Create the columns
        sColumnas = Split(.P9, "|")
        For i = LBound(sColumnas) To UBound(sColumnas)
            Set Columna = lswList.ColumnHeaders.Add(, , Separate(sColumnas(i), "^", 1), , Separate(sColumnas(i), "^", 3))
            Columna.Tag = Separate(sColumnas(i), "^", 2)
            Set Columna = Nothing
        Next i
        
        ' I create the rows
        ClearVismVars
        .P0 = USER
        Do
            .Execute ("S P1=$O(^" & sGlobal & "(P0,P1))")
            
            If .P1 = Empty Then Exit Do
            
            .Execute ("S P2=^" & sGlobal & "(P0,P1)")
            For i = LBound(sColumnas) To UBound(sColumnas)
                If i = 0 Then
                    Set Item = lswList.ListItems.Add(, , Separate(.P2, "^", i + 1))
                Else
                    Item.SubItems(i) = Separate(.P2, "^", i + 1)
                End If
            Next i
            
            Set Item = Nothing
        Loop
                
    End With
    
    ' Adjusts the width of the column to the longest content
    AutosizeColumns lswList
    
    ' We can have a fixed width
    For i = LBound(sColumnas) To UBound(sColumnas)
        If Separate(sColumnas(i), "^", 5) <> "" Then
            lswList.ColumnHeaders(i + 1).Width = Separate(sColumnas(i), "^", 5)
        End If
    Next i

    Me.Enabled = True
End Sub

Private Sub btnExport_Click()
    Dim i As Integer
    Dim Item As ListItem
    Dim bRet As Boolean
    
    If lswList.ListItems.Count = 0 Then Exit Sub
    
    bRet = False
    Select Case IniGet(PATH_INI, "CONFIG", "EXPORT", "3")
        Case "1": bRet = Export2Calc(lswList, ProgressBar)
        Case "2": bRet = Export2Excel("", lswList, ProgressBar)
        Case "3": bRet = Export2CSV(lswList, "", ProgressBar)
    End Select
    
    If bRet Then
        MsgBox "Data exported correctly.", vbInformation + vbOKOnly, App.ProductName
    Else
        MsgBox "Data was not exported due to an error", vbCritical + vbOKOnly, App.ProductName
    End If
End Sub

Private Sub Form_Activate()
    Dim sComponents() As String
    Dim sValores() As String
    Dim i As Integer
    Dim j As Integer
    Dim iIndex As Integer
    
    If sComp <> Empty Then
    
        sComponents = Split(sComp, "|")
        
        ReDim oComponents(0 To UBound(sComponents))
        For i = LBound(sComponents) To UBound(sComponents)
            
            Select Case Separate(sComponents(i), "^", 1)
                Case "TXT":
                    iIndex = txt.Count
                    
                    If iIndex = 1 And txt(iIndex - 1).Tag = Empty Then
                        iIndex = iIndex - 1
                    Else
                        Load txt(iIndex)
                    End If
                    
                    txt(iIndex).TabIndex = i
                    txt(iIndex).Width = GetWidth(Separate(sComponents(i), "^", 3))
                    txt(iIndex).Tag = Separate(sComponents(i), "^", 1)
                    txt(iIndex).Visible = True
                    
                    Set oComponents(i) = txt(iIndex)
                Case "CBO":
                    iIndex = cbo.Count
                    
                    If iIndex = 1 And cbo(iIndex - 1).Tag = Empty Then
                        iIndex = iIndex - 1
                    Else
                        Load cbo(iIndex)
                    End If
                    
                    cbo(iIndex).TabIndex = i
                    cbo(iIndex).Width = GetWidth(Separate(sComponents(i), "^", 3))
                    cbo(iIndex).Tag = Separate(sComponents(i), "^", 1)
                    cbo(iIndex).Visible = True
                    
                    sValores = Split(Separate(sComponents(i), "^", 4), ":")
                    For j = LBound(sValores) To UBound(sValores)
                        cbo(iIndex).AddItem sValores(j)
                    Next j
                    
                    Set oComponents(i) = cbo(iIndex)
                Case "DTP":
                    iIndex = dtp.Count
                    
                    If iIndex = 1 And dtp(iIndex - 1).Tag = Empty Then
                        iIndex = iIndex - 1
                    Else
                        Load dtp(iIndex)
                    End If
                    
                    dtp(iIndex).TabIndex = i
                    dtp(iIndex).Value = Date
                    dtp(iIndex).Width = GetWidth(Separate(sComponents(i), "^", 3))
                    dtp(iIndex).Tag = Separate(sComponents(i), "^", 1)
                    dtp(iIndex).Visible = True
                    
                    Set oComponents(i) = dtp(iIndex)
                Case "CHK":
                    iIndex = chk.Count
                    
                    If iIndex = 1 And chk(iIndex - 1).Tag = Empty Then
                        iIndex = iIndex - 1
                    Else
                        Load chk(iIndex)
                    End If
                    
                    chk(iIndex).TabIndex = i
                    chk(iIndex).Value = 0
                    chk(iIndex).Width = GetWidth(Separate(sComponents(i), "^", 3))
                    chk(iIndex).Tag = Separate(sComponents(i), "^", 1)
                    chk(iIndex).Visible = True
                    
                    Set oComponents(i) = chk(iIndex)
            End Select
            
            iIndex = lbl.Count
            
            If iIndex = 1 And lbl(iIndex - 1).Caption = Empty Then
                iIndex = iIndex - 1
            Else
                Load lbl(iIndex)
            End If
    
            If Separate(sComponents(i), "^", 1) = "CHK" Then
                lbl(iIndex).Caption = ""
                lbl(iIndex).Visible = False
            Else
                lbl(iIndex).Caption = Separate(sComponents(i), "^", 2)
                lbl(iIndex).Visible = True
            End If
        Next i
        
    
    End If
    
    AcomodarObjetos
    
End Sub

Private Sub AcomodarObjetos()
    Dim i As Integer
    Dim lTotal As Long
        
    If sComp <> Empty Then
        lTotal = 0
        For i = LBound(oComponents) To UBound(oComponents)
            If i = 0 Then
                oComponents(i).Left = 120
            Else
                oComponents(i).Left = oComponents(i - 1).Left + oComponents(i - 1).Width + 120
            End If
            
            lbl(i).Left = oComponents(i).Left
            lTotal = lTotal + oComponents(i).Width + 120
            
            If i = UBound(oComponents) Then
                btnSearch.Left = oComponents(i).Left + oComponents(i).Width + 120
                frFilters.Width = lTotal + btnSearch.Width + 240
                btnSearch.TabIndex = oComponents(i).TabIndex + 1
            End If
        Next i
    Else
        btnSearch.Left = 120
        btnSearch.Top = 240
        frFilters.Width = btnSearch.Width + 240
        frFilters.Height = btnSearch.Height + 360
        frList.Top = frFilters.Top + frFilters.Height
        Form_Resize
    End If
End Sub

Private Sub Form_Resize()
    If (Me.Height - frFilters.Height - btnExport.Height - 720) > 0 Then frList.Height = Me.Height - frFilters.Height - btnExport.Height - 720
    If (Me.Width - 480) > 0 Then frList.Width = Me.Width - 480
    If (frList.Height - 360) > 0 Then lswList.Height = frList.Height - 360
    If (frList.Width - 240) > 0 Then lswList.Width = frList.Width - 240
    btnExport.Left = frList.Width - btnExport.Width + 120
    btnExport.Top = frList.Top + frList.Height + 120
    
    ProgressBar.Left = frList.Left
    ProgressBar.Top = btnExport.Top
    If (Me.Height - btnExport.Width - 120) > 0 Then ProgressBar.Width = Me.Height - btnExport.Width - 120
End Sub

Private Sub lswList_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lswList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrderListView Me, lswList, ColumnHeader
End Sub

Private Sub txt_GotFocus(Index As Integer)
    PaintObj txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    EnterTab KeyAscii
    ForceCaps KeyAscii
End Sub

Private Sub txt_LostFocus(Index As Integer)
    PaintObj txt(Index), vbWhite, False
End Sub
