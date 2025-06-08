currentVersion = "1.1.2"
Public Const lokalita1 = "Praha"
Public Const lokalita1_adresa = "Vodafone náměstí Junkových 2 155 00 Praha"
Public Const lokalita2 = "Chrudim"
Public Const lokalita2_adresa = "Vodafone Průmyslová 890 537 01 Chrudim"
Public Const lokalita3 = "Ostrava"
Public Const lokalita3_adresa = "Vodafone Josefa Šavla 684/10 708 00 Ostrava"
Public Const lokalita4 = "Homeoffice"
Public Const lokalita4_adresa = ""     
Public Const input_file = "Export.xlsx"
Public Const input_sheet = "Sheet1"
Public Const output_file = "Final.csv"
Public Const output_TMP_file = "FinalTMP.csv"

' Výběr adresy podle názvu lokality ve jménu skriptu
Select Case Mid(WScript.ScriptName, InStr(WScript.ScriptName, "@") + 1, Len(WScript.ScriptName) - InStr(WScript.ScriptName, "@") - 4)
    Case lokalita1
        lokalita = lokalita1_adresa
    Case lokalita2
        lokalita = lokalita2_adresa
    Case lokalita3
        lokalita = lokalita3_adresa
    Case lokalita4
        lokalita = lokalita4_adresa
    Case Else
        WScript.Echo "Lokalita nenalezena"
        WScript.Quit
End Select

' Inicializace FileSystemObject a zjištění adresáře skriptu
Set fso = CreateObject("Scripting.FileSystemObject")
fDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Otevření Excelu a listu, kontrola existence
On Error Resume Next
Set fSheet = GetObject(fDir & "\" & input_file).Worksheets(input_sheet)
If Err.Number <> 0 Then
    WScript.Echo "List '" & input_sheet & "' nebo soubor '" & input_file & "' nenalezen"
    Err.Clear
    WScript.Quit
End If
On Error GoTo 0

' Smazání dočasného výstupního souboru, pokud existuje
If fso.FileExists(fDir & "\" & output_TMP_file) Then fso.DeleteFile(fDir & "\" & output_TMP_file)

' Vytvoření nového dočasného výstupního souboru a zápis hlavičky
Set fFile = fso.CreateTextFile(fDir & "\" & output_TMP_file, True)
fFile.WriteLine("Subject,Start Date,Start Time,End Date,End Time,All Day Event,Location,Private")

' Zpracování řádků od třetího řádku (i = 3)
i = 3
Do While fSheet.Cells(i, 1) <> ""
    ' Pokud ve sloupci 5 není "Yes"
    If fSheet.Cells(i, 5) <> "Yes" Then
        datum = fSheet.Cells(i, 3)
        ' Pokud je ve sloupci 18 "NVP-M" nebo je ve sloupci 9 cokoliv, zapíše se Volno
        If fSheet.Cells(i, 18) = "NVP-M" Or fSheet.Cells(i, 9) <> "" Then
            fFile.WriteLine("Volno (NVP-M)," & datum & ",0:00," & datum & ",0:00,True,,")
        Else
            ' Pokud je ve sloupci 13 cokoliv, zapíše se Dovolená
            If fSheet.Cells(i, 13) <> "" Then
                fFile.WriteLine("Dovolená," & datum & ",0:00," & datum & ",0:00,True,,")
            ' Pokud je ve sloupci 11 cokoliv, zapíše se SickDay
            ElseIf fSheet.Cells(i, 11) <> "" Then
                fFile.WriteLine("SickDay," & datum & ",0:00," & datum & ",0:00,True,,")
            Else
                ' Pokud je ve sloupci 7 cokoliv, zapíší se tři řádky (směna, oběd, směna)
                If fSheet.Cells(i, 7) <> "" Then
                    fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 6), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 7), 4) & ",False," & lokalita & ",")
                    fFile.WriteLine("Oběd," & datum & "," & FormatDateTime(fSheet.Cells(i, 7), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 8), 4) & ",False," & lokalita & ",")
                    fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 8), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 15), 4) & ",False," & lokalita & ",")
                ' Jinak se zapíše pouze jedna směna
                Else
                    fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 6), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 15), 4) & ",False," & lokalita & ",")
                End If
            End If
        End If
    End If
    i = i + 1
Loop

' Uvolnění objektů a převod do výsledného souboru s UTF-8
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