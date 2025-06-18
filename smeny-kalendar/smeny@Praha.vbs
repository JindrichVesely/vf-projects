' --- Verze skriptu ---
currentVersion = "1.2.1"

Function IsNewerVersion(current, remote)
    Dim curParts, remParts, i
    curParts = Split(current, ".")
    remParts = Split(remote, ".")
    For i = 0 To UBound(curParts)
        If CInt(remParts(i)) > CInt(curParts(i)) Then
            IsNewerVersion = True
            Exit Function
        ElseIf CInt(remParts(i)) < CInt(curParts(i)) Then
            IsNewerVersion = False
            Exit Function
        End If
    Next
    IsNewerVersion = False
End Function

' --- Kontrola nové verze ---
On Error Resume Next
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://raw.githubusercontent.com/JindrichVesely/vf-projects/main/smeny-kalendar/version.txt", False
http.Send
If http.Status = 200 Then
    remoteVersion = Trim(Split(http.responseText, vbLf)(0))
    If IsNewerVersion(currentVersion, remoteVersion) Then
        Set http2 = CreateObject("MSXML2.XMLHTTP")
        http2.Open "GET", "https://raw.githubusercontent.com/JindrichVesely/vf-projects/main/smeny-kalendar/changelog.txt", False
        http2.Send
        If http2.Status = 200 Then
            changelog = http2.responseText
        Else
            changelog = "(Zmeny se nepodarilo nacist)"
        End If
        MsgBox "Je dostupna novejsi verze skriptu (aktualni: " & currentVersion & ")" & vbCrLf & _
               "Nova verze: " & remoteVersion & vbCrLf & vbCrLf & _
               "Zmeny:" & vbCrLf & changelog, vbExclamation, "Aktualizace k dispozici"
        WScript.Quit
    End If
Else
    MsgBox "Nepodarilo se overit verzi na GitHubu.", vbExclamation
End If
On Error GoTo 0

Public Const lokalita1 = "Praha"
Public Const lokalita1_adresa = "Vodafone náměstí Junkových 2 155 00 Praha"
Public Const lokalita2 = "Chrudim"
Public Const lokalita2_adresa = "Vodafone Průmyslová 890 537 01 Chrudim"
Public Const lokalita3 = "Ostrava"
Public Const lokalita3_adresa = "Vodafone Josefa Šavla 684/10 708 00 Ostrava"
Public Const lokalita4 = "Homeoffice"
Public Const lokalita4_adresa = ""     
Public Const lokalita5 = "Kriz"
Public Const lokalita5_adresa = ""  
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
    Case lokalita5
        lokalita = lokalita5_adresa
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
        ' Pokud je ve sloupci 20 "NVP-M" nebo je ve sloupci 11 cokoliv, zapíše se Volno
        If fSheet.Cells(i, 20) = "NVP-M" Or fSheet.Cells(i, 11) <> "" Then
            fFile.WriteLine("Volno (NVP-M)," & datum & ",0:00," & datum & ",0:00,True,,")
        Else
            ' Pokud je ve sloupci 15 cokoliv, zapíše se Dovolená
            If fSheet.Cells(i, 15) <> "" Then
                fFile.WriteLine("Dovolená," & datum & ",0:00," & datum & ",0:00,True,,")
            ' Pokud je ve sloupci 13 cokoliv, zapíše se SickDay
            ElseIf fSheet.Cells(i, 13) <> "" Then
                fFile.WriteLine("SickDay," & datum & ",0:00," & datum & ",0:00,True,,")
            Else
            If InStr(WScript.ScriptName, "@Kriz") > 0 Then
                ' Speciální režim pro @Kriz: Směna 6:00–17:00 + Oběd 9:00–10:00
            fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 6), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 17), 4) & ",False," & lokalita & ",")
        fFile.WriteLine("Oběd," & datum & "," & FormatDateTime(fSheet.Cells(i, 9), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 10), 4) & ",False," & lokalita & ",")
            
            ElseIf fSheet.Cells(i, 9) <> "" Then
        fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 6), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 9), 4) & ",False," & lokalita & ",")
        fFile.WriteLine("Oběd," & datum & "," & FormatDateTime(fSheet.Cells(i, 9), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 10), 4) & ",False," & lokalita & ",")
        fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 10), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 17), 4) & ",False," & lokalita & ",")
    
    ' Jinak pokud je něco ve sloupci 7, je to dělená směna
    ElseIf fSheet.Cells(i, 7) <> "" Then
        fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 6), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 7), 4) & ",False," & lokalita & ",")
        fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 8), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 17), 4) & ",False," & lokalita & ",")

    ' Jinak se zapíše pouze jedna směna
    Else
        fFile.WriteLine("Směna," & datum & "," & FormatDateTime(fSheet.Cells(i, 6), 4) & "," & datum & "," & FormatDateTime(fSheet.Cells(i, 17), 4) & ",False," & lokalita & ",")
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