VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fade Any Window"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUnload 
      Caption         =   "Keep the Current state"
      Height          =   360
      Left            =   6120
      TabIndex        =   8
      ToolTipText     =   "Preserve the current state.. during unload..."
      Top             =   1770
      Width           =   2325
   End
   Begin VB.TextBox txtWinTitle 
      BackColor       =   &H00C7D6C9&
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   3555
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Text            =   " [ No window selected]"
      Top             =   2820
      Width           =   3285
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   840
      Top             =   4065
   End
   Begin MSComctlLib.Slider sldTransPar 
      Height          =   300
      Left            =   1905
      TabIndex        =   3
      Top             =   3195
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   529
      _Version        =   393216
      Max             =   255
      SelStart        =   255
      TickStyle       =   3
      Value           =   255
      TextPosition    =   1
   End
   Begin MSComctlLib.ListView lvwWindows 
      Height          =   2490
      Left            =   60
      TabIndex        =   2
      Top             =   135
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4392
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdOpaque 
      Caption         =   "&Opaque"
      Default         =   -1  'True
      Height          =   345
      Left            =   3810
      TabIndex        =   1
      Top             =   3495
      Width           =   1155
   End
   Begin MSComctlLib.ListView lvwAmount 
      Height          =   2490
      Left            =   6120
      TabIndex        =   7
      Top             =   4395
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4392
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "hWnd"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Please double-click on any window-title, shown in the listbox, to make it transparent or opaque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1065
      Left            =   5880
      TabIndex        =   9
      Top             =   555
      Width           =   2895
   End
   Begin VB.Label lblhWnd 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4020
      TabIndex        =   6
      Top             =   4260
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Window Title :"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2085
      TabIndex        =   4
      Top             =   2850
      Width           =   1260
   End
   Begin VB.Label lbllvwWindowsIndex 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4470
      TabIndex        =   0
      Top             =   4260
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IEnumWindowsSink
Dim CurrentIndex As Long


Private Sub cmdOpaque_Click()
    sldTransPar.Value = 255
    If Len(lblhWnd) <= 0 Then
        Exit Sub
    End If
    
    If CLng(lblhWnd) = 0 Then
        Exit Sub
    End If
    lvwAmount.ListItems.Item(CurrentIndex).SubItems(1) = sldTransPar.Value
    Opaque CLng(lvwWindows.ListItems.Item(CLng(lbllvwWindowsIndex)).SubItems(2))
End Sub

Private Sub cmdOpaque_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(lbllvwWindowsIndex) <= 0 Then Exit Sub
    If CLng(lbllvwWindowsIndex) = 0 Then Exit Sub
    
    lvwWindows.ListItems.Item(CLng(lbllvwWindowsIndex)).Selected = True
    lvwWindows.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    If chkUnload.Value = 1 Then Exit Sub
    For i = 1 To lvwAmount.ListItems.Count
        Opaque CLng(Trim(lvwAmount.ListItems.Item(i)))
    Next i
End Sub

Private Property Get IEnumWindowsSink_Identifier() As Long
    IEnumWindowsSink_Identifier = Me.hWnd
End Property

Private Sub Form_Load()
    
    lvwWindows.ColumnHeaders.Add , , "Title", lvwWindows.Width / 3
    lvwWindows.ColumnHeaders.Add , , "Class", lvwWindows.Width / 3
    lvwWindows.ColumnHeaders.Add , , "hWnd", lvwWindows.Width / 3

    EnumerateWindows Me
    sysWnd = FindWindow("Shell_traywnd", "")
    Dim itmX As ListItem
    Set itmX = lvwWindows.ListItems.Add(, , "Task-Bar")
    itmX.SubItems(1) = ClassName(sysWnd)
    itmX.SubItems(2) = sysWnd
End Sub

Private Sub IEnumWindowsSink_EnumWindow(ByVal hWnd As Long, bStop As Boolean)
Dim itmX As ListItem
    If Len(Trim(WindowTitle(hWnd))) > 0 And IsWindowVisible(hWnd) Then
        If (StrComp(ClassName(hWnd), "Progman") = 0) Or (StrComp(ClassName(hWnd), "msvb_lib_tooltips") = 0) Then Exit Sub
        
        Set itmX = lvwWindows.ListItems.Add(, , WindowTitle(hWnd))
        itmX.SubItems(1) = ClassName(hWnd)
        itmX.SubItems(2) = hWnd
    End If
End Sub

Private Sub lvwWindows_DblClick()
Dim Amount As Long, lvwAmountIndex As Long
    If Len(lbllvwWindowsIndex) <= 0 Then Exit Sub
    
    If CLng(lbllvwWindowsIndex) = 0 Then Exit Sub
    
    txtWinTitle = lvwWindows.ListItems.Item(CLng(lbllvwWindowsIndex))
    lblhWnd = lvwWindows.ListItems.Item(CLng(lbllvwWindowsIndex)).SubItems(2)
    
    Dim i As Long
    For i = 1 To lvwAmount.ListItems.Count
        If StrComp(lblhWnd, lvwAmount.ListItems.Item(i)) = 0 Then
            lvwAmountIndex = i
        Exit For
        End If
    Next i
    
    Dim itmX As ListItem
    
    If lvwAmountIndex > 0 Then
        CurrentIndex = lvwAmountIndex
    Else
        Set itmX = lvwAmount.ListItems.Add(1, , lblhWnd)
        itmX.SubItems(1) = "255"
        CurrentIndex = 1
    End If
    
    sldTransPar.Value = CLng(lvwAmount.ListItems.Item(CurrentIndex).SubItems(1))
    
    Me.Height = 4335
    
End Sub

Private Sub lvwWindows_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lbllvwWindowsIndex = lvwWindows.SelectedItem.Index
End Sub

Private Sub sldTransPar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(lblhWnd) <= 0 Then
        sldTransPar.Value = 255
        Exit Sub
    End If
    
    If CLng(lblhWnd) = 0 Then
        sldTransPar.Value = 255
        Exit Sub
    End If
    Transparent CLng(lblhWnd), sldTransPar.Value
    lvwAmount.ListItems.Item(CurrentIndex).SubItems(1) = sldTransPar.Value
    lvwWindows.ListItems.Item(CLng(lbllvwWindowsIndex)).Selected = True
    lvwWindows.SetFocus
End Sub

Private Sub sldTransPar_Scroll()
    If Len(lblhWnd) <= 0 Then
        sldTransPar.Value = 255
        Exit Sub
    End If
    
    If CLng(lblhWnd) = 0 Then
        sldTransPar.Value = 255
        Exit Sub
    End If
End Sub

Private Sub Timer1_Timer()
    If lvwWindows.ListItems.Count Then
        lvwWindows.ListItems.Clear
    End If
    
    EnumerateWindows Me
    Dim itmX As ListItem
    Set itmX = lvwWindows.ListItems.Add(, , "Task-Bar")
    itmX.SubItems(1) = ClassName(sysWnd)
    itmX.SubItems(2) = sysWnd
    
    Dim i As Long, j As Long
    Dim Found As Boolean, StillOpen As Boolean
    
    For i = 1 To lvwWindows.ListItems.Count
        If StrComp(lvwWindows.ListItems.Item(i).SubItems(2), lblhWnd) = 0 Then
            lvwWindows.ListItems.Item(i).Selected = True
            lbllvwWindowsIndex = i
            Found = True
            Exit For
        End If
    Next i
    If Not Found Then
        sldTransPar.Value = 255
        lblhWnd = 0
        txtWinTitle = " [ No window selected]"
    End If
    i = 1
    Do While i <= lvwAmount.ListItems.Count
        For j = 1 To lvwWindows.ListItems.Count
            StillOpen = False
            If StrComp(lvwWindows.ListItems.Item(j).SubItems(2), lvwAmount.ListItems.Item(i)) = 0 Then
                StillOpen = True
                Exit For
            End If
        Next j
        If Not StillOpen Then
            lvwAmount.ListItems.Remove (i)
        Else
            i = i + 1
        End If
   Loop
End Sub
