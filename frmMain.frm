VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hex Convert"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVB 
      Height          =   285
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1410
      Width           =   3195
   End
   Begin VB.TextBox txtC 
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   840
      Width           =   3195
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert!"
      Default         =   -1  'True
      Height          =   375
      Left            =   810
      TabIndex        =   0
      Top             =   1890
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000D&
      Caption         =   "Converts C Hex values (Example: 0x0001) to Visual Basic Hex values (Example: &&H1)"
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB Hex:"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1170
      Width           =   585
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter C Hex:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   630
      Width           =   2820
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'
'I have improved the function thanks to Paul Mather
'I guess I need to improve my knowings in strings :-)
'
'**Now the function is 9 lines instead of 30 ... :-)
'*****************************************************

'Application to convert C Hex values to Visual Basic
'By Max Raskin, March 2000
'
'Comments and Bugfixes: maxim13@internet-zahav.net

'********************************************************************************
'Example: CHexToVBHex("0x001L")
'Operations: 1. Zeros are removed from the middle until there is non-zero number
'            2. 0x converted to &H
'            3. If in the end of the string there is l or L its converted to &
'Result: &H1L
'********************************************************************************

Private Sub Form_Activate()
    txtC.SetFocus
End Sub

Private Sub cmdConvert_Click()
    txtVB.Text = CHexToVBHex(txtC.Text)
End Sub

'CHexToVBHex version 2.00
Function CHexToVBHex(CHex As String) As String
On Error GoTo ErrHandler
    Dim TmpStr As String, L As Boolean
    If LCase(Right(CHex, 1)) = "l" Then L = True
    TmpStr = "&H" & Hex(Val("&H" & Replace(UCase(Mid(CHex, 3)), "L", "&")))
    If L = True Then
        CHexToVBHex = Trim(TmpStr) & "&"
        L = False
    Else
        CHexToVBHex = Trim(TmpStr)
    End If
ErrHandler: Exit Function
End Function
