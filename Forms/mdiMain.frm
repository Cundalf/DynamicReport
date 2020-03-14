VERSION 5.00
Object = "{88F75480-0574-11D0-8085-0000C0BD354B}#1.0#0"; "VISM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "CADR"
   ClientHeight    =   6855
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10185
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VISMLib.VisM VisM 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      PLIST           =   ""
      PDELIM          =   $"mdiMain.frx":058A
      Interval        =   1000
      P0              =   ""
      P1              =   ""
      P2              =   ""
      P3              =   ""
      P4              =   ""
      P5              =   ""
      P6              =   ""
      P7              =   ""
      P8              =   ""
      P9              =   ""
      VALUE           =   ""
      Code            =   ""
      NameSpace       =   ""
      TimeOut         =   0
      ExecFlag        =   0
   End
   Begin MSComctlLib.StatusBar sbBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6540
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnMain 
      Caption         =   "Menu1"
      Index           =   1
      Begin VB.Menu mnSubMain1 
         Caption         =   "Tip"
         Index           =   1
      End
   End
   Begin VB.Menu mnMain 
      Caption         =   "Menu2"
      Index           =   2
      Begin VB.Menu mnSubMain2 
         Caption         =   ""
         Index           =   1
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "About"
      Index           =   3
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()
    AutoGenerateMenu
End Sub

Private Sub MDIForm_Load()
    VisM.MServer = IniGet(PATH_INI, "DB", "IP")
    VisM.NameSpace = IniGet(PATH_INI, "DB", "NAMESPACE")

    sbBar.Panels(1).Text = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub AutoGenerateMenu()
    Dim i As Integer

    ClearVismVars
    With VisM
        .P0 = USER
        .P1 = IDAPPMENU
        .Execute ("D GETMENUS^CADR(P0,P1)")

        If .error > 0 Then
            MsgBox "There was an error loading the menus." & vbCrLf & .ErrorName, vbCritical + vbOKOnly, Me.Caption
            Exit Sub
        End If

        ClearVismVars
        .P0 = USER
        Do
            .Execute ("S P1=$O(^CADR(P0,P1))")

            If .P1 = Empty Then Exit Do

            ' Clear all dynamic menu
            For i = IndexToMenu(.P1).Count To 1 Step -1
                If IndexToMenu(.P1)(i).Tag <> "" Then
                    If i = 1 Then
                        IndexToMenu(.P1)(i).Caption = ""
                        IndexToMenu(.P1)(i).Tag = ""
                    Else
                        Unload IndexToMenu(.P1)(i)
                    End If
                End If
            Next i
            
            ' Add new dynamic menu
            Do
                .Execute ("S P2=$O(^CADR(P0,P1,P2))")

                If .P2 = Empty Then Exit Do

                .Execute ("S P3=^CADR(P0,P1,P2)")

                Call AddSubMenu(Separate(.P3, "^", 1), IndexToMenu(.P1), .P2)
            Loop
        Loop

    End With
End Sub

Private Sub mnMenu_Click(Index As Integer)
    Select Case Index
        Case 3: frmAbout.Show
    End Select
End Sub

Private Sub mnSubMain1_Click(Index As Integer)

    Select Case Index
        Case 1: frmTip.Show ' Example Form
    End Select

    AutoGenerateForm mnSubMain1, Index
End Sub

Private Sub mnSubMain2_Click(Index As Integer)
    AutoGenerateForm mnSubMain2, Index
End Sub
