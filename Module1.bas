Attribute VB_Name = "Module1"
Option Explicit

Sub Button1_Click()
  Dim Sh As Worksheet, R As Integer, C As Integer, TableRows As New Collection, Txt As String, ColWidths(100) As Integer, CurrCell As Range, CellText As String
  Set Sh = ActiveSheet
  
  For R = 1 To Sh.UsedRange.Rows.Count
    For C = 1 To Sh.UsedRange.Columns.Count
      If Len(Sh.Cells(R, C).Text) > ColWidths(C) Then ColWidths(C) = Len(Sh.Cells(R, C).Text)
    Next C
  Next R
  
  TableRows.Add "{| class=""wikitable sortable"" style=""text-align: center;"""
  For R = 1 To Sh.UsedRange.Rows.Count
    TableRows.Add "|-"
    Txt = "| "
    For C = 1 To Sh.UsedRange.Columns.Count
      Set CurrCell = Sh.Cells(R, C)
      CellText = Replace(CurrCell.Text, "|", vbTab)
      CellText = Replace(CellText, vbLf, "<br>")
      If CurrCell.Font.Color = vbRed Then CellText = "nowrap|" & CellText
      If CurrCell.MergeArea.Columns.Count > 1 And CurrCell.MergeArea.Rows.Count > 1 Then
        If CurrCell.Column = CurrCell.MergeArea.Column And CurrCell.Row = CurrCell.MergeArea.Row Then Txt = Txt & "colspan=""" & CurrCell.MergeArea.Columns.Count & """ rowspan=""" & CurrCell.MergeArea.Rows.Count & """" & vbTab & CellText & " || "
      ElseIf CurrCell.MergeArea.Columns.Count > 1 Then
        If CurrCell.Column = CurrCell.MergeArea.Column Then Txt = Txt & "colspan=""" & CurrCell.MergeArea.Columns.Count & """" & vbTab & CellText & " || "
      ElseIf CurrCell.MergeArea.Rows.Count > 1 Then
        If CurrCell.Row = CurrCell.MergeArea.Row Then Txt = Txt & "rowspan=""" & CurrCell.MergeArea.Rows.Count & """" & vbTab & CellText & " || "
      Else
        Txt = Txt & ExtendString(CellText, ColWidths(C)) & " || "
      End If
    Next C
    If Cells(R, 1).Font.Bold Then Txt = Replace(Txt, "|", "!")
    Txt = Replace(Txt, vbTab, "|")
    Txt = Replace(Txt, "|nowrap|", " nowrap|")
    Txt = Replace(Txt, "|width=", " width=")
    TableRows.Add Left(Txt, Len(Txt) - 4)
  Next R
  TableRows.Add "|}"
  
  Txt = ""
  For R = 1 To TableRows.Count
    Txt = Txt & TableRows(R) & vbCrLf
  Next R
  
  PutInClipboard2 Txt
End Sub

Function ExtendString(Txt As String, Length As Integer) As String
  ExtendString = Txt & String(Length - Len(Txt), " ")
End Function

Sub PutInClipboard(Txt As String)
  Dim DataObj As New MSForms.DataObject
  DataObj.SetText Txt
  DataObj.PutInClipboard
End Sub

Sub PutInClipboard2(Txt As String)
  Dim TempFile As String
  TempFile = Environ("Temp") & "\WikiTable.txt"
  Open TempFile For Output As #1
  Print #1, Txt;
  Close #1
  
  Dim wsh As Object
  Set wsh = VBA.CreateObject("WScript.Shell")
  wsh.Run "cmd /c clip < " & TempFile, 1, True
  
  Kill TempFile
End Sub
