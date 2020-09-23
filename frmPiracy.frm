VERSION 5.00
Begin VB.Form frmPiracy 
   Caption         =   "Let's Go Pirate Hunting!!"
   ClientHeight    =   4455
   ClientLeft      =   2385
   ClientTop       =   1545
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5805
   Begin Piracy.ctlPiracy ctlPiracy 
      Height          =   675
      Left            =   3930
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin VB.CommandButton cmdSaveName 
      Caption         =   "&Save This Registration Name"
      Height          =   375
      Left            =   450
      TabIndex        =   2
      Top             =   3810
      Width           =   2535
   End
   Begin VB.TextBox txtRegName 
      Height          =   285
      Left            =   480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   3450
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "scumbag"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "cheapskate"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   2730
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "pennypincher"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   2430
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Cheater"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   2130
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Test with these Illegal Names:"
      Height          =   255
      Index           =   0
      Left            =   510
      TabIndex        =   5
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "This Control should be visible=false"
      Height          =   285
      Left            =   3030
      TabIndex        =   4
      Top             =   1410
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   $"frmPiracy.frx":0000
      Height          =   975
      Left            =   150
      TabIndex        =   1
      Top             =   90
      Width           =   5115
   End
End
Attribute VB_Name = "frmPiracy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Dim INIPath As String
Dim RegName As String
Dim Title As String
Dim Prompt As String
Dim Ret As Long



Private Sub SelectWholeText(InTextBox As Control)

  InTextBox.SelStart = 0
  InTextBox.SelLength = Len(InTextBox.Text)
End Sub

Private Sub CenterForm(frm As Form)
    Dim X, Y                    ' New top, left for the form
    X = (Screen.Width - frm.Width) / 2
    Y = (Screen.Height - frm.Height) / 2
    frm.Move X, Y             ' Change location of the form
End Sub

Private Sub cmdSaveName_Click()
 Dim I As Integer
 Dim K%
 Dim TmpStr As String
 Dim UserName As String
 
'Since we use the input# function to read the text file from the server we
'don't want to allow names with double quotes ""

UserName = Trim$(txtRegName)
TmpStr = ""
For I = 1 To Len(UserName)
 If Asc(Mid(UserName, I, 1)) = 34 Then
    TmpStr = TmpStr & " " 'Put in a space
 Else
    TmpStr = TmpStr & Mid(UserName, I, 1)
 End If
Next I
RegName = TmpStr
K% = WritePrivateProfileString("Registration", "Name", RegName, INIPath$)
MsgBox "Thank you for paying for a registration code!" & vbCrLf & vbCrLf & "You Better Hope So!", vbInformation, "Thank You"
 
End Sub

Private Sub ctlPiracy_BadConnection(Number As Integer, Description As String)
 'Don't do anything here. We don't want the user to know that we're
 'checking for illegal codes.
End Sub

Private Sub GetIniPath()
  Dim sPath As String
  
  sPath = App.Path
  'Append trailing backslash if necessary
  If Asc(Right$(sPath, 1)) <> 92 Then
    sPath = sPath & "\"
  End If
  INIPath = sPath & "Pirate.ini"

End Sub

Private Sub Form_Load()
 Dim I%
 Dim TmpStr As String
 
 GetIniPath
 
 'IMPORTANT: Always Redim PirName so when you exit program, you don't get
 'a runtime for attempting to access an undimensioned array.
 ReDim PirName(0)
 
 CenterForm Me
 
 'Let's Hope their connected
 ctlPiracy.GetProgState
 
 'Get the saved registration name:
  RegName = Space$(35)
  I% = GetPrivateProfileString("Registration", "Name", "Unregistered Copy", RegName, 34, INIPath$)
  RegName = Left$(RegName, I%)
  txtRegName = RegName
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim J As Integer
    
    For J = 0 To UBound(PirName)
     If PirName(J) = RegName Then
        'Found an Illegal Registration Name
        GiveThemTheBadNews
        Exit For
     End If
   Next J

End Sub


Private Sub txtRegName_GotFocus()
 SelectWholeText txtRegName
 
End Sub

Private Sub GiveThemTheBadNews()
  Dim I%

'   1) Write a nasty name as their registration name:
    I% = WritePrivateProfileString("Registration", "Name", "YOU THIEF", INIPath$)

'   2) Give them a stern warning:
       MsgBox "CAUGHT YOU !!", vbCritical, "You Thief !!"
   
   
   'Depending on your anger, add code to destroy their registry (system.dat)
   'or simply delete any vital data files your program might create etc...
   'Be Inventive :) You just caught one!!!
   
End Sub
