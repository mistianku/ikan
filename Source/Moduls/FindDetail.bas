Attribute VB_Name = "FindDetail"
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim kodeUserAkses As String
Public Function FindDataDetail(sKodeUserAkses As String, sTable As String, sField As String, sFieldCari As String, sConnectku As Integer) As String
On Error GoTo errhandler

    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select " & sFieldCari & " from " & sTable & " where " & sField & " ='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
       FindDataDetail = oRs(sFieldCari)
    Else
        FindDataDetail = ""
    End If
    oCon.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Function
Public Function FindDataDetailAlternatif(sKodeUserAkses As String, sTable As String, sField As String, sFieldCari As String, sDatabaseku As Integer) As String
On Error GoTo errhandler

    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select " & sFieldCari & " from " & sTable & " where " & sField & " ='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
       FindDataDetailAlternatif = oRs(sFieldCari)
    Else
        FindDataDetailAlternatif = ""
    End If
    oCon.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Function




Public Function FindFirstLast(sTable As String, sFieldFind As String, sFlagFind As FlagFind, sConnectku As Integer) As String
On Error GoTo errhandler

    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    If sFlagFind = First Then
        sQuery = "Select Top 1 " & sFieldFind & " from " & sTable & " Order By " & sFieldFind & " Asc"
    Else
        sQuery = "Select Top 1 " & sFieldFind & " from " & sTable & " Order By " & sFieldFind & " Desc"
    End If
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
       FindFirstLast = oRs(sFieldFind)
    Else
        FindFirstLast = ""
    End If
    oCon.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Function



Public Function GetDocnum(sDocumentNo As DocumentNo, sUpdateNo As Boolean, sConnectku As Integer) As String
On Error GoTo errhandler

Dim sConKu As New ADODB.Connection
Dim sRstku As New ADODB.Recordset
Dim sQuery As String
Dim sGetDocnum As String
Dim sdocnum As String
Dim sFormatdoc As String
Dim txtPPCari As String
Dim txtYYCari As String
Dim txtMMari As String
Dim txt99Cari As String
'Dim sConnectku  As Integer
If sConKu.State = 1 Then sConKu.Close
sConKu.Open MainModule.Conectionku(DBaseConection.Modul)

sQuery = "SELECT fget_docnum(" & sDocumentNo & "," & Year(Now()) & "," & Month(Now()) & ")"
Set sRstku = sConKu.Execute(sQuery)
GetDocnum = sRstku(0)
'sQuery = "select  autonodefault,docid, allowprefix, textprefix,  allowyop, allowmop, "
'sQuery = sQuery & "doclength, docnum,docnumfmt,textprefix_last,yop_last,mop_last from setting_document "
'sQuery = sQuery & " where docid=" & sDocumentNo
'Set sRstku = sConKu.Execute(sQuery)
'
'sFormatdoc = ToText(sRstku("docnumfmt"))
'
'
'
'txt99Cari = Mid(sFormatdoc, InStr(sFormatdoc, "9"), InStrRev(sFormatdoc, "9") - InStr(sFormatdoc, "9") + 1)
'If sRstku("allowprefix") = "1" Then
'    txtPPCari = Mid(sFormatdoc, InStr(sFormatdoc, "P"), InStrRev(sFormatdoc, "P") - InStr(sFormatdoc, "P") + 1)
'    sGetDocnum = Replace(sFormatdoc, txtPPCari, Trim(sRstku("textprefix")))
'Else
'    sGetDocnum = ""
'End If
'
'    If sRstku("allowyop") = "1" Then
'        txtYYCari = Mid(sFormatdoc, InStr(sFormatdoc, "Y"), InStrRev(sFormatdoc, "Y") - InStr(sFormatdoc, "Y") + 1)
'        sGetDocnum = Replace(sGetDocnum, txtYYCari, Right(Year(Now()), 2))
'    Else
'        sGetDocnum = Replace(sGetDocnum, "YY", "",dbaseConection.Modul)
'    End If
'
'            If sRstku("allowmop") = "1" Then
'                txtMMari = Mid(sFormatdoc, InStr(sFormatdoc, "M"), InStrRev(sFormatdoc, "M") - InStr(sFormatdoc, "M") + 1)
'                sGetDocnum = Replace(sGetDocnum, txtMMari, IIf(Month(Now()) < 10, "0" & Month(Now()), Month(Now())))
'            Else
'                sGetDocnum = Replace(sGetDocnum, "MM", "",dbaseConection.Modul)
'            End If
'            Dim i As Integer
'            sdocnum = ""
'            For i = 1 To sRstku("doclength")
'                sdocnum = sdocnum & "0"
'            Next
'            sdocnum = sdocnum & sRstku("docnum")
'            sGetDocnum = Replace(sGetDocnum, txt99Cari, Right(sdocnum, Len(txt99Cari) + (sRstku("doclength") - Len(sGetDocnum))))
'            'GetDocnum = sGetDocnum & Right(sdocnum, sRstku("doclength") - Len(sGetDocnum))
'            GetDocnum = sGetDocnum
'If sUpdateNo = True Then
'    sQuery = "Update setting_document set Docnum=Docnum+1 Where DocID=" & sDocumentNo
'    Set sRstku = sConKu.Execute(sQuery)
'End If
'select @DoCnum=DoCnum, @DocWith=DocWidth,@Prefix=case when Prefix is null then '' else Prefix end from PosSettingDocNumber where DocID=@DocID
sConKu.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "GetDocnum"

End Function

Public Function GetAllowAutoNumber(sModul As Modul, sConnectku As Integer, sCheck As CheckBox, sText As TextBox) As String
On Error GoTo errhandler

Dim sConKu As New ADODB.Connection
Dim sRstku As New ADODB.Recordset
Dim sQuery As String
'Dim sConnectku  As Integer
If sConKu.State = 1 Then sConKu.Close
sConKu.Open MainModule.Conectionku(DBaseConection.Modul)
sQuery = "Select AllowManlDocNum from PosSettingDocNumber Where DocID='" & sModul & "'"
Set sRstku = sConKu.Execute(sQuery)

GetAllowAutoNumber = sRstku(0)
If GetAllowAutoNumber = "0" Then
    sCheck.Enabled = False
    sText.Enabled = False
Else
    sCheck.Enabled = True
    sText.Enabled = True
End If
sConKu.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "GetDocnum"

End Function
