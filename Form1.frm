VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Long File Name to Short File Name"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Any File"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Short File Name"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Long File Name"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

   Public Function GetShortName(ByVal sLongFileName As String) As String
       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
   End Function

Private Sub Command1_Click()
Dim ShortName As String
Screen.MousePointer = 11

CommonDialog1.CancelError = True
On Error GoTo EH1

CommonDialog1.Filter = "all files (*.*)|*.*"
CommonDialog1.Flags = &H80000 Or &H1000
CommonDialog1.ShowOpen

Text1.Text = CommonDialog1.filename
ShortName = GetShortName(CommonDialog1.filename)
Text2.Text = ShortName

Screen.MousePointer = 0
Exit Sub

EH1:

Screen.MousePointer = 0
If Err = 32755 Then Err.Clear: Exit Sub
MsgBox Err.Description, vbExclamation, "ERR #" & Err
End Sub
