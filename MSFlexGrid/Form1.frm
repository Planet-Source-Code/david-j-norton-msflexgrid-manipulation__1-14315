VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MSFlexGrid - Add Rows with Alternate Colors with Option to Change+Grid Line On/Off Toggle"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Exit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   4905
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton ChangeAltRowColor 
      Caption         =   "&Change Alternate Row Color"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   4905
      Width           =   2175
   End
   Begin VB.CommandButton ToggleGridLines 
      Caption         =   "&Toggle Grid Lines On/Off"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   4905
      Width           =   2175
   End
   Begin VB.CommandButton AddRow 
      Caption         =   "&Add Row"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4905
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   4470
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4305
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   7594
         _Version        =   393216
         Cols            =   7
         BackColor       =   16777215
         BackColorFixed  =   -2147483637
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   1
         BorderStyle     =   0
      End
   End
   Begin VB.Label TotalRows 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label CurCell 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim FixedColCaptions(0 To 6) As String, C As Long, i
    
    FixedColCaptions(0) = "ID:"
    FixedColCaptions(1) = "First Name:"
    FixedColCaptions(2) = "Last Name:"
    FixedColCaptions(3) = "Address:"
    FixedColCaptions(4) = "City:"
    FixedColCaptions(5) = "State:"
    FixedColCaptions(6) = "Zip Code:"
    
    MSFlexGrid1.Row = 0
    For i = 0 To 6
        MSFlexGrid1.Col = i
        If i = 0 Then
            MSFlexGrid1.ColWidth(i) = 350
        Else
            MSFlexGrid1.ColWidth(i) = 1330
        End If
        MSFlexGrid1.Text = FixedColCaptions(i)
    Next
    
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Text = Format(MSFlexGrid1.Row, "000")
    
    GridAltColor = &HC0FFC0
    
    For i = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = GridAltColor
    Next
    
    MSFlexGrid1.Col = 1
    GridLineToggle = 1
    TotalRows.Caption = "Total Rows= " & MSFlexGrid1.Rows - 1
    
End Sub

Private Sub MSFlexGrid1_EnterCell()
    Dim Msg
    Msg = " Active Cell: " & Chr(64 + MSFlexGrid1.Col) & MSFlexGrid1.Row
    CurCell = Msg + " - " + MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Col) + " "
    MSFlexGrid1.Tag = MSFlexGrid1
End Sub

Private Sub AddRow_Click()
     Dim i As Integer
    
    MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
    GridColorCount = GridColorCount + 1
    
    If GridColorCount > 2 Then GridColorCount = 1
    
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.Row = MSFlexGrid1.Row + 1
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Text = Format(MSFlexGrid1.Row, "000")
    
    For i = 1 To MSFlexGrid1.Cols - 1
        If GridColorCount = 1 Then
            MSFlexGrid1.Col = i
            MSFlexGrid1.CellBackColor = &HFFFFFF
        Else
            MSFlexGrid1.Col = i
            MSFlexGrid1.CellBackColor = GridAltColor '&HC0FFC0
        End If
    Next
    TotalRows.Caption = "Total Rows= " & MSFlexGrid1.Rows - 1
    MSFlexGrid1.Col = 1
    MSFlexGrid1.SetFocus
End Sub

Private Sub ChangeAltRowColor_Click()
    Dim C As Long
    Dim Response
    
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    
    CommonDialog1.Flags = cdlCCRGBInit
    CommonDialog1.ShowColor
    GridAltColor = CommonDialog1.Color
    
    Response = MsgBox("Confirm Change of Grid Alternate Row Color", vbOKCancel + vbQuestion, "Options - Change Grid Colors")
    If Response = vbCancel Then GoTo ErrHandler
    
    MSFlexGrid1.Row = 0
    
    Do Until MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
        For C = 1 To MSFlexGrid1.Cols - 1
            MSFlexGrid1.Col = C
            MSFlexGrid1.CellBackColor = GridAltColor
        Next
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
    Loop
    MSFlexGrid1.Col = 1
    MSFlexGrid1.SetFocus
    Exit Sub
    
ErrHandler:
    MSFlexGrid1.Col = 1
    MSFlexGrid1.SetFocus
    Exit Sub
End Sub

Private Sub ToggleGridLines_Click()
    If GridLineToggle = 1 Then
        MSFlexGrid1.GridLines = flexGridNone
        GridLineToggle = 2
    ElseIf GridLineToggle = 2 Then
        MSFlexGrid1.GridLines = flexGridFlat
        GridLineToggle = 1
    End If
    MSFlexGrid1.SetFocus
End Sub

Private Sub Exit_Click()
    Dim Msg
    
    Msg = "Please Vote For Me!" + Chr$(10)
    Msg = Msg + "N.B.  Some of the Code was contributed by Raul Lopez and Modified by me. " + Chr$(10) + Chr$(10)
    Msg = Msg + "Thanks" + Chr$(10)
    Msg = Msg + "David J Norton"
    MsgBox Msg, vbOKOnly + vbInformation, "MSFlexGrid Sample Code"
    
    End
    
End Sub


