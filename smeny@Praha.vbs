Public Const lokalita1 = "Praha"
Public Const lokalita1_adresa = "Vodafone náměstí Junkových 2 155 00 Praha"
Public Const lokalita2 = "Chrudim"
Public Const lokalita2_adresa = "Vodafone Průmyslová 890 537 01 Chrudim"
Public Const lokalita3 = "Ostrava"
Public Const lokalita3_adresa = "Vodafone Josefa Šavla 684/10 708 00 Ostrava"
Public Const input_file = "Export.xlsx"
Public Const input_sheet = "Sheet1"
Public Const output_file = "Final.csv"
Public Const output_TMP_file = "FinalTMP.csv"

Select Case Mid(WScript.ScriptName, InStr(WScript.ScriptName, "@") + 1, Len(WScript.ScriptName) - InStr(WScript.ScriptName, "@") - 4)
    Case lokalita1
        lokalita = lokalita1_adresa
    Case lokalita2
        lokalita = lokalita2_adresa
    Case lokalita3_adresa
        lokalita = lokalita3_adresa
    Case Else
        WScript.Echo "Lokalita nenalezena"
        WScript.Quit
End Select
Set fso = CreateObject("Scripting.FileSystemObject")
fDir = fso.GetParentFolderName(WScript.ScriptFullName)
On Error Resume Next
Set fSheet = GetObject(fDir & "\" & input_file).Worksheets(input_sheet)
If Err.Number <> 0 Then
    WScript.Echo "List '" & input_sheet & "' nebo soubor '" & input_file & "' nenalezen"
    Err.Clear
    WScript.Quit
End If
On Error GoTo 0
If fso.FileExists(fDir & "\" & output_TMP_file) Then fso.DeleteFile(fDir & "\" & output_TMP_file)
Set fFile = fso.CreateTextFile(fDir & "\" & output_TMP_file, True)
fFile.WriteLine("Subject,Start Date,Start Time,End Date,End Time,All Day Event,Location,Private")
i = 3
Do While fSheet.Cells(i, 1) <> ""
    If fSheet.Cells(i, 5) <> "Yes" Then 'Date - Nothing flag - preskakuji Yes
        datum = fSheet.Cells(i, 3)
        Select Case fSheet.Cells(i, 12)
            Case "NVP-M"
                fFile.WriteLine("Volno," & datum & ",0:00," & datum & ",0:00,True,,")
            Case "VACA"
                fFile.WriteLine("Dovolená," & datum & ",0:00," & datum & ",0:00,True,,")
            Case Else
                If fSheet.Cells(i, 7) <> "" Then
                    fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 6), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 7), 4) & ",False," & lokalita & ",")
                    fFile.WriteLine("Oběd," & datum & "," & FormatDateTime(fSheet.Cells(i, 7), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 8), 4) & ",False," & lokalita & ",")
                    fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 8), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 9), 4) & ",False," & lokalita & ",")
                Else
                    fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 6), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 9), 4) & ",False," & lokalita & ",")
                End If
        End Select
    End If
    i = i + 1
Loop
Set fSheet = Nothing
fFile.Close
Set stream = CreateObject("ADODB.Stream")
stream.Open
stream.Type = 2
stream.Charset = "utf-8"
stream.LoadFromFile fDir & "\" & output_TMP_file
text = stream.ReadText
stream.Close
If fso.FileExists(fDir & "\" & output_file) Then fso.DeleteFile(fDir & "\" & output_file)
Set fFile = fso.OpenTextFile(fDir & "\" & output_file, 2, True, True)
fFile.Write text
fFile.Close
Set fFile = Nothing
If fso.FileExists(fDir & "\" & output_TMP_file) Then fso.DeleteFile(fDir & "\" & output_TMP_file)
Set fso = Nothing
WScript.Echo "Hotovo"
WScript.Quit
