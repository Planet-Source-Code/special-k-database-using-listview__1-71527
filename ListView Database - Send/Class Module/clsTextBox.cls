VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Requried for text box masking
'-------------------------------------------------------------------

Option Explicit
Private oWorking As Boolean

Public Sub MaskedText(TextObj As TextBox, ByVal MyProfile As String)

 Dim MyCursorPlace As Long
 Dim MyLen As Long
 Dim MyPlace As Long
 Dim MyBuffer As String
 Dim MyText As String
 Dim MyChar As String * 1
 Dim MyProfilePlace As Long
 Const GoodMatchColor = vbBlue
 Const BadMatchColor = vbRed
  
  
    If oWorking Then Exit Sub
    oWorking = True
 
    On Local Error Resume Next

    MyPlace = 1
    MyProfilePlace = 1
    
    'If there are more characters than allowed then remove them
    
    If Len(TextObj.Text) > Len(MyProfile) Then
        
        TextObj.Text = Left(TextObj.Text, Len(MyProfile))
        TextObj.SelStart = Len(TextObj.Text)
        
    End If
    
    MyText = TextObj.Text
    MyLen = Len(MyText)
    MyCursorPlace = TextObj.SelStart
    
    ' The parser takes the pattern as the transition map. Starting
    ' at the beginning of the map, it compares the current character
    ' with the state of the parser
    
    Do While MyPlace <= MyLen
        
        MyChar = Mid(MyText, MyPlace, 1)
        Select Case Mid(MyProfile, MyProfilePlace, 1)
            ' Check if MyProfile calls for a numeric input */
         Case "#"
            If IsNumeric(MyChar) Then
                ' Add the character to the buffer
                MyBuffer = MyBuffer & MyChar
                ' Move to the next character
                MyPlace = MyPlace + 1
                ' Move to the next valid parser state
                MyProfilePlace = MyProfilePlace + 1
                ' Make sure we are indicating a valid transition state
                TextObj.ForeColor = GoodMatchColor
            Else
                ' The character does not match the parser's state so
                ' indicate an invalid state and exit the parser
                TextObj.ForeColor = BadMatchColor
                
                GoTo ExitCheckProfile
            End If
         Case "@"
            If Not IsNumeric(MyChar) Then
                ' Add the character to the buffer
                MyBuffer = MyBuffer & MyChar
                ' Move to the next character
                MyPlace = MyPlace + 1
                ' Move to the next valid parser state
                MyProfilePlace = MyProfilePlace + 1
                ' Make sure we are indicating a valid transition state
                TextObj.ForeColor = GoodMatchColor
            Else
                ' The character does not match the parser's state so
                ' indicate an invalid state and exit the parser
                TextObj.ForeColor = BadMatchColor
                
                GoTo ExitCheckProfile
            End If
         Case Else
            If MyChar = Mid(MyProfile, MyProfilePlace, 1) Then
                ' If it Is here Then add the character to the buffer
                MyBuffer = MyBuffer & MyChar
                ' Move to next character position
                MyPlace = MyPlace + 1
                ' Move to next parser state */
                MyProfilePlace = MyProfilePlace + 1
                ' Indicate a valid transition state to the user
                TextObj.ForeColor = GoodMatchColor
            Else
                ' The required character is not present and in this case we
                ' insert it meeting the requirements of the parser state
                MyBuffer = MyBuffer & Mid(MyProfile, MyProfilePlace, 1)
                ' Shift the parser to the next state but stay with the current
                ' character to see if it matches the next state transition
                MyProfilePlace = MyProfilePlace + 1
                ' Move the input cursor over one to account for the added character
                MyCursorPlace = MyCursorPlace + 1
            End If
        End Select
    Loop
    
    TextObj.Text = MyBuffer
    TextObj.SelStart = MyCursorPlace
    
ExitCheckProfile:
    oWorking = False

End Sub


' Requried for text box masking Ends
'-------------------------------------------------------------------

