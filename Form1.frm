VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Dictionary Search"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Height          =   885
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3240
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "firewall"
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2160
      Picture         =   "Form1.frx":1272
      Top             =   300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''
''''''SubZero DeZignS'''''''''
''''''''''''''''''''''''''''''

Option Explicit
Dim Search As String
Dim Second As String
Dim spot As Integer
Dim spot2 As Integer
Dim done As String, done2 As String
Dim beginspot As Integer, EndSpot As Integer

Sub GetString(word As String)
On Error Resume Next
' Here is where we open the connection to the said site with Inet control
Text1.Text = Trim(Text1.Text)
Inet1.URL = "http://encarta.msn.com/dictionary_/" & Text1.Text & ".html"
' Now we dump the text of the said site into the Text box
Text = Inet1.OpenURL(Inet1.URL)
' Now we call upon our first string parser
GetDefinition
End Sub

Sub GetDefinition()
'We start by going down the rows for the parsing.
On Error GoTo Err
Search = "CORE MEANING:"
spot = InStr(Text, "CORE MEANING:")
done = Mid(Text, spot - 2, spot)
spot2 = InStr(done, "<br />")
done = Mid(done, 1, spot2 - 1)
'Replace Begin
done = pReplace(Trim(done), "CORE MEANING:", "CORE MEANING: ")
done = pReplace(Trim(done), ">", "")
done = Trim(done)
'Parser Finish 'Replace Finish
'Dump into Label for Viewing.
Label1 = done
Exit Sub
Err:
' If that parser was not quite hitting the spot we call upon the next string parser
Alternate
Exit Sub
End Sub

Sub Alternate()
'As you see I named this one Alternative. This one will get a string that closely matches your word match
On Error GoTo Err
spot = InStr(Text, "but we found the following alternate spellings for you.")
done = Mid(Text, spot, spot)
spot2 = InStr(done, "Search all of MSN Encarta for")
done = Mid(done, 1, spot2 - 1)
'Replace Begin
done = pReplace(Trim(done), "but we found the following alternate spellings for you.", "Alternate: ")
done = pReplace(Trim(done), "Click one to continue your search.", "")
done = pReplace(Trim(done), "</td>", "")
done = pReplace(Trim(done), "</tr>", "")
done = pReplace(Trim(done), """", "")
done = pReplace(Trim(done), "<tr height=", "")
done = pReplace(Trim(done), "16>", "")
done = pReplace(Trim(done), "<td height=", "")
done = pReplace(Trim(done), "<tr><td class=", "")
done = pReplace(Trim(done), "class=", "")
done = pReplace(Trim(done), "NoResultsSuggestions", "")
done = pReplace(Trim(done), "<a href=", "")
done = pReplace(Trim(done), "/dictionary_/", "")
done = pReplace(Trim(done), """", "")
done = pReplace(Trim(done), ".html>", " ")
done = pReplace(Trim(done), "</a>", " ")
done = pReplace(Trim(done), "10>", "")
done = pReplace(Trim(done), "10>", "")
done = pReplace(Trim(done), "SearchAll>", "")
done = pReplace(Trim(done), ">", "")
done = pReplace(Trim(done), "", "")
done = pReplace(Trim(done), "", "")
done = Trim(done)
'Parser Finish 'Replace Finish
'Dump into Label for Viewing.
Label1 = done
Exit Sub
Err:
Alternative2
Exit Sub
End Sub

Sub Alternative2()
'This one is for strings that contain more then one definition
'If you want. you could put more into the parser like spot = InStr(Text, "3.&nbsp;") and start another string.
'Remembering to change the String qualifier. "done, done2, done3" ext'
On Error GoTo Err
spot = InStr(Text, ">1.&nbsp;")
done = Mid(Text, spot, spot)
spot2 = InStr(done, "<br />")
done = Mid(done, 1, spot2 - 1)
done = pReplace(Trim(done), "", "")
done = pReplace(Trim(done), "&nbsp;", " ")
done = pReplace(Trim(done), "</b>", "")
done = pReplace(Trim(done), "<b>", "")
done = pReplace(Trim(done), "<span", "")
done = pReplace(Trim(done), "ResultBody", "")
done = pReplace(Trim(done), "SmallCaps", "")
done = pReplace(Trim(done), "</span>", "")
done = pReplace(Trim(done), "(", "")
done = pReplace(Trim(done), ")", "")
done = pReplace(Trim(done), "</td>", "")
done = pReplace(Trim(done), "</tr>", "")
done = pReplace(Trim(done), """", "")
done = pReplace(Trim(done), "<tr height=", "")
done = pReplace(Trim(done), "16>", "")
done = pReplace(Trim(done), "<td height=", "")
done = pReplace(Trim(done), "<tr><td class=", "")
done = pReplace(Trim(done), "class=", "")
done = pReplace(Trim(done), "NoResultsSuggestions", "")
done = pReplace(Trim(done), "<a href=", "")
done = pReplace(Trim(done), "/dictionary_/", "")
done = pReplace(Trim(done), """", "")
done = pReplace(Trim(done), ".html>", " ")
done = pReplace(Trim(done), "</a>", " ")
done = pReplace(Trim(done), "10>", "")
done = pReplace(Trim(done), "10>", "")
done = pReplace(Trim(done), "SearchAll>", "")
done = pReplace(Trim(done), "<i>", "")
done = pReplace(Trim(done), "</i>", "")
done = pReplace(Trim(done), vbCrLf, " ")
done = pReplace(Trim(done), ">", "")
done = pReplace(Trim(done), Chr(10), " ")
done = pReplace(Trim(done), Chr(13), " ")
done = pReplace(Trim(done), Chr(13), "")
done = pReplace(Trim(done), Chr(13), "")
done = pReplace(Trim(done), Chr(13), "")
done = pReplace(Trim(done), Chr(13), "")
done = pReplace(Trim(done), Chr(10), "")
done = pReplace(Trim(done), Chr(10), "")
done = pReplace(Trim(done), Chr(10), "")

'Second Definitions start here.
spot = InStr(spot, Text, ">2.&nbsp;")
done2 = Mid(Text, spot, spot)
spot2 = InStr(done2, "<br />")
done2 = Mid(done2, 1, spot2 - 1)
done2 = pReplace(Trim(done2), "", "")
done2 = pReplace(Trim(done2), "&nbsp;", " ")
done2 = pReplace(Trim(done2), "</b>", "")
done2 = pReplace(Trim(done2), "<b>", "")
done2 = pReplace(Trim(done2), "<span", "")
done2 = pReplace(Trim(done2), "ResultBody", "")
done2 = pReplace(Trim(done2), "SmallCaps", "")
done2 = pReplace(Trim(done2), "</span>", "")
done2 = pReplace(Trim(done2), "(", "")
done2 = pReplace(Trim(done2), ")", "")
done2 = pReplace(Trim(done2), "</td>", "")
done2 = pReplace(Trim(done2), "</tr>", "")
done2 = pReplace(Trim(done2), """", "")
done2 = pReplace(Trim(done2), "<tr height=", "")
done2 = pReplace(Trim(done2), "16>", "")
done2 = pReplace(Trim(done2), "<td height=", "")
done2 = pReplace(Trim(done2), "<tr><td class=", "")
done2 = pReplace(Trim(done2), "class=", "")
done2 = pReplace(Trim(done2), "NoResultsSuggestions", "")
done2 = pReplace(Trim(done2), "<a href=", "")
done2 = pReplace(Trim(done2), "/dictionary_/", "")
done2 = pReplace(Trim(done2), """", "")
done2 = pReplace(Trim(done2), ".html>", " ")
done2 = pReplace(Trim(done2), "</a>", " ")
done2 = pReplace(Trim(done2), "10>", "")
done2 = pReplace(Trim(done2), "10>", "")
done2 = pReplace(Trim(done2), "SearchAll>", "")
done2 = pReplace(Trim(done2), "<i>", "")
done2 = pReplace(Trim(done2), "</i>", "")
done2 = pReplace(Trim(done2), ">", "")
done2 = Trim(done2)
Label1 = done & vbCrLf & done2
Exit Sub
Err:
Alternative3
Exit Sub
End Sub

Sub Alternative3()
'This ones for such language MSN doesn't allow.
On Error GoTo Err
'Looks for the Word string. if it returns Language Advisory
'Then I +'ed it 17 for putting that word string back into the label for visual conformation.
spot = InStr(Text, "Language Advisory")
done = Mid(Text, spot, spot)
spot2 = InStr(done, "Language Advisory") + 17
done = Mid(done, 1, spot2 - 1)
done = pReplace(Trim(done), "", "")
done = Trim(done)
Label1 = done
Exit Sub
Err:
'And if all else fails. There's probably no more definitions to parse for the word
Label1 = "No Matches found. Retry another Word or shorten it."
Exit Sub
End Sub

Private Sub Form_Load()
Form1.Height = 2800
Form1.Width = 3800
End Sub

Private Sub Form_Resize()
Text1.Width = Form1.Width
Label1.Width = Form1.Width
Label1.Height = Form1.Height - Text1.Height
End Sub

Private Sub Label1_Change()
If Label1.Caption = "Language Advisory" Then Image1.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'If they pressed the Enter/Return key
If KeyAscii = vbKeyReturn Then
KeyAscii = 0
Inet1.Cancel
Label1.Caption = "Connecting. -"
Image1.Visible = False
Text.Text = ""
GetString Text1.Text
done = pReplace(Trim(done), "", "")
End If
End Sub
