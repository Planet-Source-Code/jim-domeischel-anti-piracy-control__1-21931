VERSION 5.00
Begin VB.UserControl ctlPiracy 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ctlPiracy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Event BadConnection(Number As Integer, Description As String)

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
Dim TempStr As String
Dim I As Integer
Dim FNum As Integer
Dim TempFName As String

FNum = FreeFile

On Error GoTo DLoadError
    If AsyncProp.PropertyName = "Suzy" Then
        TempFName = AsyncProp.Value
        I = 0
        Open TempFName For Input As FNum
        Do While Not EOF(FNum)
           ReDim Preserve PirName(I)
           Input #FNum, TempStr
           PirName(I) = Trim$(TempStr)
           I = I + 1
        Loop
    End If
    Exit Sub
    
DLoadError:
    RaiseEvent BadConnection(1028, "")

End Sub

Public Sub GetProgState()
Dim HURL As String
On Error GoTo HeadersError
    
    'Place a simple text file on your server and fill it with the illegal
    'registration names in Input # format (read the  VB documentation for Input#).
    'i.e "Cheater","pennypincher","cheapskate" etc...
    'You can test it (and get caught) with the following registration names:
     
'     Cheater
'     pennypincher
'     cheapskate
'     scumbag
    
    'If they put out another crack, just add it to the text file :)
    'SEE CRACKS.TXT
    
    HURL = "http://www.jimsquest.com/cracks.txt"
    AsyncRead HURL, vbAsyncTypeFile, "Suzy"
    Exit Sub

HeadersError:
    RaiseEvent BadConnection(1024, "")

End Sub

