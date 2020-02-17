VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExcelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Class file for writing Microsoft Excel BIFF 2.1 files.

'This class is intended for users who do not want to use the huge
'Jet or ADO providers if they only want to export their data to
'an Excel compatible file.

'Newer versions of Excel use the OLE Structure Storage methods
'which are quite complicated.

'Paul Squires, October 3, 2000
'rambo2000@canada.com


'enum to handle the various types of values that can be written
'to the excel file.
Public Enum ValueTypes
    xlsInteger = 0
    xlsNumber = 1
    xlsText = 2
End Enum

'enum to hold cell alignment
Public Enum CellAlignment
    xlsGeneralAlign = 0
    xlsLeftAlign = 1
    xlsCentreAlign = 2
    xlsRightAlign = 3
    xlsFillCell = 4
    xlsLeftBorder = 8
    xlsRightBorder = 16
    xlsTopBorder = 32
    xlsBottomBorder = 64
    xlsShaded = 128
End Enum

'enum to handle selecting the font for the cell
Public Enum CellFont
    'used by rgbAttr2
    'bits 0-5 handle the *picture* formatting, not bold/underline etc...
    'bits 6-7 handle the font number
    xlsFont0 = 0
    xlsFont1 = 64
    xlsFont2 = 128
    xlsFont3 = 192
End Enum

Public Enum CellHiddenLocked
    'used by rgbAttr1
    'bits 0-5 must be zero
    'bit 6 locked/unlocked
    'bit 7 hidden/not hidden
    xlsNormal = 0
    xlsLocked = 64
    xlsHidden = 128
End Enum


'set up variables to hold the spreadsheet's layout
Public Enum MarginTypes
   xlsLeftMargin = 38
   xlsRightMargin = 39
   xlsTopMargin = 40
   xlsBottomMargin = 41
End Enum


Public Enum FontFormatting
   'add these enums together. For example: xlsBold + xlsUnderline
   xlsNoFormat = 0
   xlsBold = 1
   xlsItalic = 2
   xlsUnderline = 4
   xlsStrikeout = 8
End Enum


Private Type FONT_RECORD
   opcode As Integer  '49
   length As Integer  '5+len(fontname)
   FontHeight As Integer
   
   'bit0 bold, bit1 italic, bit2 underline, bit3 strikeout, bit4-7 reserved
   FontAttributes1 As Byte
   
   FontAttributes2 As Byte  'reserved - always 0
   
   FontNameLength As Byte
End Type


Private Type PASSWORD_RECORD
   opcode As Integer  '47
   length As Integer  'len(password)
End Type


Private Type HEADER_FOOTER_RECORD
   opcode As Integer  '20 Header, 21 Footer
   length As Integer  '1+len(text)
   TextLength As Byte
End Type


Private Type PROTECT_SPREADSHEET_RECORD
   opcode As Integer  '18
   length As Integer  '2
   Protect As Integer
End Type

Private Type COLWIDTH_RECORD
   opcode As Integer  '36
   length As Integer  '4
   col1 As Byte       'first column
   col2 As Byte       'last column
   ColumnWidth As Integer   'at 1/256th of a character
End Type

'Beginning Of File record
Private Type BEG_FILE_RECORD
  opcode As Integer
  length As Integer
  version As Integer
  ftype As Integer
End Type

'End Of File record
Private Type END_FILE_RECORD
  opcode As Integer
  length As Integer
End Type

'true/false to print gridlines
Private Type PRINT_GRIDLINES_RECORD
  opcode As Integer
  length As Integer
  PrintFlag As Integer
End Type

'Integer record
Private Type tInteger
  opcode As Integer
  length As Integer
  row As Integer     'unsigned integer
  col As Integer
  
  'rgbAttr1 handles whether cell is hidden and/or locked
  rgbAttr1 As Byte
  
  'rgbAttr2 handles the Font# and Formatting assigned to this cell
  rgbAttr2 As Byte
  
  'rgbAttr3 handles the Cell Alignment/borders/shading
  rgbAttr3 As Byte
  
  intValue As Integer  'the actual integer value
End Type

'Number record
Private Type tNumber
  opcode As Integer
  length As Integer
  row As Integer
  col As Integer
  rgbAttr1 As Byte
  rgbAttr2 As Byte
  rgbAttr3 As Byte
  NumberValue As Double  '8 Bytes
End Type

'Label (Text) record
Private Type tText
  opcode As Integer
  length As Integer
  row As Integer
  col As Integer
  rgbAttr1 As Byte
  rgbAttr2 As Byte
  rgbAttr3 As Byte
  TextLength As Byte
End Type

Private Type MARGIN_RECORD_LAYOUT
  opcode As Integer
  length As Integer
  MarginValue As Double  '8 bytes
End Type


Dim FileNumber As Integer
Dim BEG_FILE_MARKER As BEG_FILE_RECORD
Dim END_FILE_MARKER As END_FILE_RECORD



Public Function CreateFile(ByVal FileName As String) As Integer

On Error GoTo Write_Error

    If Dir$(FileName) > "" Then
       Kill FileName
    End If
    
    FileNumber = FreeFile
    Open FileName For Binary As #FileNumber
    Put #FileNumber, , BEG_FILE_MARKER  'must always be written first
    
    OpenFile = 0  'return with no error
    
Exit Function

Write_Error:
    OpenFile = Err.Number
    Exit Function

End Function

Public Function CloseFile() As Integer

On Error GoTo Write_Error

    If FileNumber = 0 Then Exit Function
    
    Put #FileNumber, , END_FILE_MARKER
    Close #FileNumber

    CloseFile = 0  'return with no error code
    
Exit Function

Write_Error:
    CloseFile = Err.Number
    Exit Function

End Function


Private Sub Class_Initialize()

'Set up default values for records
'These should be the values that are the same for every record of these types
    
    With BEG_FILE_MARKER  'beginning of file
        .opcode = 9
        .length = 4
        .version = 2
        .ftype = 10
    End With
    
    With END_FILE_MARKER  'end of file marker
        .opcode = 10
    End With
    
    
End Sub


Public Function WriteValue(ValueType As ValueTypes, CellFontUsed As CellFont, Alignment As CellAlignment, HiddenLocked As CellHiddenLocked, lrow As Long, lcol As Long, value As Variant) As Integer

On Error GoTo Write_Error

'the row and column values are written to the excel file as
'unsigned integers. Therefore, must convert the longs to integer.
    If lrow > 32767 Then
       row% = CInt(lrow - 65536) - 1  'rows/cols in Excel binary file are zero based
    Else
       row% = CInt(lrow) - 1
    End If
    
    If lcol > 32767 Then
       col% = CInt(lcol - 65536) - 1  'rows/cols in Excel binary file are zero based
    Else
       col% = CInt(lcol) - 1
    End If

    
    Select Case ValueType
      Case ValueTypes.xlsInteger
         Dim INTEGER_RECORD As tInteger
         With INTEGER_RECORD
           .opcode = 2
           .length = 9
           .row = row%
           .col = col%
           .rgbAttr1 = CByte(HiddenLocked)
           .rgbAttr2 = CByte(CellFontUsed)
           .rgbAttr3 = CByte(Alignment)
           .intValue = CInt(value)
         End With
         Put #FileNumber, , INTEGER_RECORD
    
    
      Case ValueTypes.xlsNumber
         Dim NUMBER_RECORD As tNumber
         With NUMBER_RECORD
           .opcode = 3
           .length = 15
           .row = row%
           .col = col%
           .rgbAttr1 = CByte(HiddenLocked)
           .rgbAttr2 = CByte(CellFontUsed)
           .rgbAttr3 = CByte(Alignment)
           .NumberValue = CDbl(value)
         End With
         Put #FileNumber, , NUMBER_RECORD
      
      
      Case ValueTypes.xlsText
         Dim b As Byte
         If IsNull(value) Then value = ""
         st$ = CStr(value)
         l% = Len(st$)
        
        Dim TEXT_RECORD As tText
        With TEXT_RECORD
          .opcode = 4
          .length = 10
          'Length of the text portion of the record
          .TextLength = l%
        
          'Total length of the record
          .length = 8 + l
        
          .row = row%
          .col = col%
          
          .rgbAttr1 = CByte(HiddenLocked)
          .rgbAttr2 = CByte(CellFontUsed)
          .rgbAttr3 = CByte(Alignment)
        
          'Put record header
          Put #FileNumber, , TEXT_RECORD
        
          'Then the actual string data
          For a = 1 To l%
             b = Asc(Mid$(st$, a, 1))
             Put #FileNumber, , b
          Next
        End With
      
    End Select
    
    WriteValue = 0   'return with no error
    
Exit Function

Write_Error:
    WriteValue = Err.Number
    Exit Function

End Function


Public Function SetMargin(Margin As MarginTypes, MarginValue As Double) As Integer

On Error GoTo Write_Error

    'write the spreadsheet's layout information (in inches)
    Dim MarginRecord As MARGIN_RECORD_LAYOUT
    
    With MarginRecord
      .opcode = Margin
      .length = 8
      .MarginValue = MarginValue 'in inches
    End With
    Put #FileNumber, , MarginRecord

    SetMargin = 0
    
Exit Function

Write_Error:
    SetMargin = Err.Number
    Exit Function

End Function


Public Function SetColumnWidth(FirstColumn As Byte, LastColumn As Byte, WidthValue As Integer)

On Error GoTo Write_Error

    Dim COLWIDTH As COLWIDTH_RECORD
    
    With COLWIDTH
      .opcode = 36
      .length = 4
      .col1 = FirstColumn - 1
      .col2 = LastColumn - 1
      .ColumnWidth = WidthValue * 256  'values are specified as 1/256 of a character
    End With
    Put #FileNumber, , COLWIDTH

    SetColumnWidth = 0
    
Exit Function

Write_Error:
    SetColumnWidth = Err.Number
    Exit Function

End Function


Public Function SetFont(FontName As String, FontHeight As Integer, FontFormat As FontFormatting) As Integer

On Error GoTo Write_Error

    'you can set up to 4 fonts in the spreadsheet file. When writing a value such
    'as a Text or Number you can specify one of the 4 fonts (numbered 0 to 3)
    
    Dim FONTNAME_RECORD As FONT_RECORD
    
    l% = Len(FontName)
    
    With FONTNAME_RECORD
      .opcode = 49
      .length = 5 + l%
      .FontHeight = FontHeight * 20
      .FontAttributes1 = CByte(FontFormat)  'bold/underline etc...
      .FontAttributes2 = CByte(0) 'reserved-always zero!!
      .FontNameLength = CByte(Len(FontName))
    End With
    Put #FileNumber, , FONTNAME_RECORD

    'Then the actual font name data
    Dim b As Byte
    For a = 1 To l%
       b = Asc(Mid$(FontName, a, 1))
       Put #FileNumber, , b
    Next

    SetFont = 0
    
Exit Function

Write_Error:
    SetFont = Err.Number
    Exit Function


End Function


Public Function SetHeader(HeaderText As String) As Integer

On Error GoTo Write_Error

    Dim HEADER_RECORD As HEADER_FOOTER_RECORD
    
    l% = Len(HeaderText)
    
    With HEADER_RECORD
      .opcode = 20
      .length = 1 + l%
      .TextLength = CByte(Len(HeaderText))
    End With
    Put #FileNumber, , HEADER_RECORD

    'Then the actual Header text
    Dim b As Byte
    For a = 1 To l%
       b = Asc(Mid$(HeaderText, a, 1))
       Put #FileNumber, , b
    Next

    SetHeader = 0
    
Exit Function

Write_Error:
    SetHeader = Err.Number
    Exit Function

End Function



Public Function SetFooter(FooterText As String) As Integer

On Error GoTo Write_Error

    Dim FOOTER_RECORD As HEADER_FOOTER_RECORD
    
    l% = Len(FooterText)
    
    With FOOTER_RECORD
      .opcode = 21
      .length = 1 + l%
      .TextLength = CByte(Len(FooterText))
    End With
    Put #FileNumber, , FOOTER_RECORD

    'Then the actual Header text
    Dim b As Byte
    For a = 1 To l%
       b = Asc(Mid$(FooterText, a, 1))
       Put #FileNumber, , b
    Next

    SetFooter = 0
    
Exit Function

Write_Error:
    SetFooter = Err.Number
    Exit Function

End Function



Public Function SetFilePassword(PasswordText As String) As Integer

On Error GoTo Write_Error

    Dim FILE_PASSWORD_RECORD As PASSWORD_RECORD
    
    l% = Len(PasswordText)
    
    With FILE_PASSWORD_RECORD
      .opcode = 47
      .length = l%
    End With
    Put #FileNumber, , FILE_PASSWORD_RECORD

    'Then the actual Password text
    Dim b As Byte
    For a = 1 To l%
       b = Asc(Mid$(PasswordText, a, 1))
       Put #FileNumber, , b
    Next

    SetFilePassword = 0
    
Exit Function

Write_Error:
    SetFilePassword = Err.Number
    Exit Function

End Function




Public Property Let PrintGridLines(ByVal newvalue As Boolean)

On Error GoTo Write_Error

    Dim GRIDLINES_RECORD As PRINT_GRIDLINES_RECORD
    
    With GRIDLINES_RECORD
      .opcode = 43
      .length = 2
      If newvalue = True Then
        .PrintFlag = 1
      Else
        .PrintFlag = 0
      End If
      
    End With
    Put #FileNumber, , GRIDLINES_RECORD

Exit Property

Write_Error:
    Exit Property

  
End Property




Public Property Let ProtectSpreadsheet(ByVal newvalue As Boolean)

On Error GoTo Write_Error

    Dim PROTECT_RECORD As PROTECT_SPREADSHEET_RECORD
    
    With PROTECT_RECORD
      .opcode = 18
      .length = 2
      If newvalue = True Then
        .Protect = 1
      Else
        .Protect = 0
      End If
      
    End With
    Put #FileNumber, , PROTECT_RECORD

Exit Property

Write_Error:
    Exit Property

  
End Property




