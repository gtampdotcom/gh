Attribute VB_Name = "modColor"
'This loads the first custom colors row with predetermined colors
'and loads the second custom color row with user selected custom colors
'and saves the choices in a file.
'The only thing I have not figured out is how to set the color box focus
'it will always load r = 0 g = 0 b = 0
'If you can figure this out email me at sdowney@erols.com


Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim clr As Long

Private Type CHOOSECOLOR
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   rgbResult As Long
   lpCustColors As String
   Flags As Long
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias _
   "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Dim CustomColors() As Byte

Public Sub modNumberColor(ColorNumber, BoolCancel) ' colornumber is the color chosen, boolcancel user pressed the cancel button
    
    On Error Resume Next ' used for when file does not exist
    ReDim CustomColors(0 To 63) As Byte
      
           'first row
           'box1
           CustomColors(0) = 139   'red
           CustomColors(1) = 155   'green
           CustomColors(2) = 184   'blue
           'box2
           CustomColors(4) = 188   'red
           CustomColors(5) = 213   'green
           CustomColors(6) = 254   'blue
           'box3
           CustomColors(8) = 115   'red
           CustomColors(9) = 172   'green
           CustomColors(10) = 183  'blue
           'box4
           CustomColors(12) = 200    'red
           CustomColors(13) = 249   'green
           CustomColors(14) = 198   'blue
           'box5
           CustomColors(16) = 189   'red
           CustomColors(17) = 194   'green
           CustomColors(18) = 253   'blue
           'box6
           CustomColors(20) = 200    'red
           CustomColors(21) = 249   'green
           CustomColors(22) = 255   'blue
           'box7
           CustomColors(24) = 108    'red
           CustomColors(25) = 213   'green
           CustomColors(26) = 210   'blue
           'box8
           CustomColors(28) = 236   'red
           CustomColors(29) = 164   'green
           CustomColors(30) = 236   'blue
           
           'default colors load
           '2nd row
           'box9
           CustomColors(32) = 160   'red
           CustomColors(33) = 160  'green
           CustomColors(34) = 160  'blue
           '10
           CustomColors(36) = 160   'red
           CustomColors(37) = 160  'green
           CustomColors(38) = 160  'blue
           '11
           CustomColors(40) = 160   'red
           CustomColors(41) = 160  'green
           CustomColors(42) = 160  'blue
           '12
           CustomColors(44) = 160   'red
           CustomColors(45) = 160  'green
           CustomColors(46) = 160  'blue
           '13
           CustomColors(48) = 160   'red
           CustomColors(49) = 160  'green
           CustomColors(50) = 160  'blue
           '14
           CustomColors(52) = 160   'red
           CustomColors(53) = 160  'green
           CustomColors(54) = 160  'blue
           '15
           CustomColors(56) = 160   'red
           CustomColors(57) = 160  'green
           CustomColors(58) = 160  'blue
           '16
           CustomColors(60) = 160  'red
           CustomColors(61) = 160  'green
           CustomColors(62) = 160  'blue
       
       'read custom colors out of file and load into custom color array
'       FileNum = FreeFile
'       Open ((App.Path + "\CustomColors.ini")) For Input As #FileNum
'       For i = 32 To 62
'       Input #FileNum, i, CustomColors(i)
'       Next i
'       Close #FileNum
       
       Dim cc As CHOOSECOLOR
       'Dim Custcolor(16) As Long
       Dim lReturn As Long
       
       ' Store the initial settings of the Choose Color box.
            
       cc.lStructSize = Len(cc) ' size of the structure

       ' If you comment out the following line,
       ' the dialog appears in the upper left corner
       ' of the screen.
       cc.hwndOwner = Form1.hwnd 'Form1 is opening the Choose Color box

       cc.hInstance = 0 'not needed
       'cc.rgbResult = frmLogon.BackColor  'doesnt work 'set default selected color to Form1's background color
       cc.lpCustColors = StrConv(CustomColors, vbUnicode)
       cc.Flags = CC_ANYCOLOR Or CC_RGBINIT 'allow any color, use rgbResult as default selection
       cc.lCustData = 0  ' not needed
       cc.lpfnHook = 0  ' not needed
       cc.lpTemplateName = ""  ' not needed
       
       'open color dialog box
       lReturn = ChooseColorAPI(cc)
       
       If lReturn <> 0 Then
           ColorNumber = cc.rgbResult
           CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
           
           'store custom colors in a file
'           FileNum = FreeFile
'           Open ((App.Path + "\CustomColors.ini")) For Output As #FileNum
'           For i = 32 To 63
'           Write #FileNum, i, CustomColors(i)
'           Next i
'           Close #FileNum
       Else
           BoolCancel = True
           Exit Sub
       End If
  
End Sub
   
Public Function ColorView() As String
 
  modNumberColor ColorNumber, BoolCancel 'module2
  If BoolCancel = True Then Exit Function
  'Form1.BackColor = ColorNumber
  clr = ColorNumber
  'unRGB clr, r, g, b
  ColorView = clr
  
  'Label1.Caption = "Color: " & Format$(clr)
  'Label2.Caption = "Red: " & Format$(r) & " Green: " & Format$(g) & " Blue: " & Format$(b)
  'Me.Caption = "Color: " & Format$(clr) & "    Red: " & Format$(r) & " Green: " & Format$(g) & " Blue: " & Format$(b)
End Function
'Private Sub unRGB(ByVal clr As Long, r As Integer, g As Integer, b As Integer)
'
'  r = clr Mod 256
'  g = (clr \ 256) Mod 256
'  b = clr \ 256 \ 256
'
'End Sub
