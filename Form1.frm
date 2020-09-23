VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Find String in Textbox Example"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Find &Last"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Case Sensitive Search"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Text            =   "strawberry"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find &Next"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Find First"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Search for"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str As String
Dim i As Integer
Dim n As Integer

Private Sub Command1_Click()
'find first occurence
i = InStr(1, str, Text2.Text, IIf(Check1.Value, vbBinaryCompare, vbTextCompare))
If i Then
   Label2.Caption = "Found at position: " & CStr(i)
   Command2.Enabled = True
   i = i - 1
   Text1.SetFocus
   Text1.SelStart = i
   Text1.SelLength = Len(Text2.Text)
   'Note: after selecting the text you can retrieve it via
   'Text1.SelText
   i = i + Len(Text2.Text) + 1
   n = 1
Else
   Command2.Enabled = False
   MsgBox "Not found!"
End If
End Sub

Private Sub Command2_Click()
'find first occurence
i = InStr(i, str, Text2.Text, IIf(Check1.Value, vbBinaryCompare, vbTextCompare))
If i Then
   Label2.Caption = "Found at position: " & CStr(i)
   i = i - 1
   Text1.SetFocus
   Text1.SelStart = i
   Text1.SelLength = Len(Text2.Text)
   i = i + Len(Text2.Text) + 1
   n = n + 1
Else
   Command2.Enabled = False
   MsgBox CStr(n) & " found!"
End If
End Sub

Private Sub Command3_Click()
'find last occurence
i = InStrRev(str, Text2.Text, -1, IIf(Check1.Value, vbBinaryCompare, vbTextCompare))
If i Then
   Label2.Caption = "Found at position: " & CStr(i)
   Command2.Enabled = False
   i = i - 1
   Text1.SetFocus
   Text1.SelStart = i
   Text1.SelLength = Len(Text2.Text)
   i = i + Len(Text2.Text) + 1
   n = 1
Else
   MsgBox "Not found!"
End If
End Sub

Private Sub Form_Load()
Text1.Text = "Let me take you down cuz I'm going to Strawberry Fields...nothing is real...and nothing to get hung about. Strawberry Fields forever. " & _
  "Living is easy with eyes closed...misunderstanding all you see...it's getting hard to be someone but it all works out...it doesn't matter much to me. " & _
  "Let me take you down cuz I'm going to Strawberry Fields...nothing is real...and nothing to get hung about. Strawberry Fields forever...Strawberry Fields forever..." & _
  "Strawberry Fields forever!"
str = Text1.Text
End Sub
