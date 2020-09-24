VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Pig Latin Text Conversion"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Text Converter"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert Text to Pig Latin"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      X1              =   120
      X2              =   4440
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000013&
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Caption         =   "Enter a phrase or word to convert to Pig Latin."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------
' Program: Pig Latin Text Conversion
' Objective: Convert sentence to Pig Latin format
' Author: Bradley Buskey
' Date: February 29, 2000
'-------------------------------------------------------

Option Explicit

Private Sub Form_Load()
'   Here we clear anything out of the text box.
    txtSource.Text = ""
End Sub

Private Sub cmdConvert_Click()
'   Variable Declaration
    Dim Source() As String
    Dim MsgTitle, MsgType, Temp, Hold, TextNum, FullText
    Dim n As Integer
    
'   Set up the Title and Type of message box.
    MsgTitle = "Your Converted Message"
    MsgType = vbInformation
    
'   Bring in the entire phrase into an array split on a space.
    Source() = Split(txtSource.Text, " ")
    
'   Loop through from 0 to the final entry in the array.
    For n = 0 To UBound(Source())
    
'       Get the number of characters in the word.
        TextNum = Len(Source(n))
        
'       Take the first character and save it to a temporary variable.
        Temp = Left(Source(n), 1)
        
'       Take all the characters less the first on into a temporary variable.
        Hold = Right(Source(n), TextNum - 1)
        
'       Remake the word, pig latin-style.
        Source(n) = Hold & Temp & "ay"
        
'       Put the sentence back together.
        FullText = FullText & " " & Source(n)
    
'   Go to the next word.
    Next n
    
'   Display the re-done sentence in pig latin style.
    MsgBox FullText, MsgType, MsgTitle
End Sub

Private Sub txtSource_DblClick()
'   This is so it is easy to clear the textbox.
    txtSource.Text = ""
End Sub

Private Sub cmdExit_Click()
'   Pretty self explanitory, eh?
    End
End Sub
