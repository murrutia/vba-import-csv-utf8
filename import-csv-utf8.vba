Option Explicit

Sub import()
    
    ' Boîte de dialogue pour choisir le fichier à charger
    Dim file_path As String: file_path = Application.GetOpenFilename("Fichiers CSV, *.csv")
    'If file_path = False Then Exit Sub
    Dim file_name As String: file_name = GetFilenameFromPath(file_path)
        
    ' Pour éviter d'excéder la mémoire d'excel lors de la lecture de fichiers entiers
    ' (ce qui est fait plus tard par LoadUTF8File()) on découpe le fichier toutes les 5000 lignes
    Dim nb_files As Integer: nb_files = SplitFile(file_path, 5000)
    
    ' Ouverture d'un Workbook où seront rajoutés les résultats
    Dim wb As Workbook: Set wb = Workbooks.Add
    
    Dim ws As Worksheet
    Dim i As Integer
    Dim file_content As String
    
    For i = 1 To nb_files
    
        ' Chargement du contenu du fichier avec l'encoding UTF-8  (pour que les accents
        ' s'affichent correctement) et suppresion du fichier temporaire
        file_content = LoadUTF8FileAndDelete(file_path & "_" & i & ".csv")
        
        Set ws = GetWorksheetForFilenameAndNum(wb, file_name, i)
        
        ' Découpage du contenu du fichier en "morceaux" de CSV et injection dans la Worksheet
        call InjectCsvChunksIntoWorksheet(GetCsvChunks(file_content), ws)
                            
    Next

    ' Suppression de la première Worksheet (vide) du nouveau Workbook et focus sur celle
    ' contenant les premiers résultats 
    Application.DisplayAlerts = False
    wb.Sheets(1).Delete
    Application.DisplayAlerts = True
    wb.Sheets(1).Activate

    MsgBox "Traitement terminé"
End Sub

' source : https://stackoverflow.com/a/1743356
Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

' Découpe le fichier passé en paramètre en fonction du nombre de lignes passé en paramètre
' et renvoie le nombre de fichiers résultant
Function SplitFile(ByVal file_path As String, ByVal ws_lines_limit As Integer) As Integer
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim txt_stream As TextStream: Set txt_stream = fso.OpenTextFile(file_path, ForReading, False)
    Dim line_nb As Integer: line_nb = 1
    Dim nb_files As Integer: nb_files = 1
    
    ' La première ligne d'un fichier CSV contient en général les headers de colonnes,
    ' donc on la sauvegarde pour l'insérer dans chaque fichier
    Dim first_line As String: first_line = txt_stream.ReadLine
    
    Dim o_file As Object: Set o_file = fso.CreateTextFile(file_path & "_" & nb_files & ".csv")
    
    o_file.WriteLine first_line
    
    Do While Not txt_stream.AtEndOfStream
        ' A chaque fois que la limite de lignes est atteinte dans une Worksheet,
        ' on ferme le fichier courant et on en réouvre un nouveau 
        If line_nb = ws_lines_limit Then
            o_file.Close
            nb_files = nb_files + 1
            Set o_file = fso.CreateTextFile(file_path & "_" & nb_files & ".csv")
            o_file.WriteLine first_line
            line_nb = 1
        End If

        line_nb = line_nb + 1
        o_file.WriteLine txt_stream.ReadLine
    Loop
    
    o_file.Close
    Set fso = Nothing
    Set o_file = Nothing
    txt_stream.Close
    
    SplitFile = nb_files
End Function

Function LoadUTF8FileAndDelete(ByVal file_path As String) As String
    Dim File As ADODB.Stream
    Set File = CreateObject("ADODB.Stream")
    File.Open
    File.Type = 2 'adTypeText
    File.Charset = "UTF-8"
    File.LoadFromFile file_path
    LoadUTF8FileAndDelete = File.ReadText
    File.Close
    Kill file_path
End Function

Function GetWorksheetForFilenameAndNum(ByRef wb As Workbook, ByVal file_name As String, ByVal num As Integer) As Worksheet
    Dim ws As Worksheet

    ' J'en fais un peu trop ici en récupérant une feuille si elle existe déjà et la créant sinon.
    ' Maintenant il me suffirait de la créer parce que je travaille toujours avec un nouveau Workbok,
    ' mais au début du développement je réutilisais toujours le même Workbook.
    ' Je garde ce code uniquement pour avoir un (mauvais?) exemple de gestion d'erreur dans VBA
    On Error Resume Next
    Set ws = Nothing
    Set ws = wb.Sheets(file_name & " " & num)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = file_name & " " & num
    End If
    Set getWorksheetForFilenameAndNum = ws
End Function

Function GetCsvChunks(ByRef file_content As String) As VBScript_RegExp_55.MatchCollection
    Dim reg As New VBScript_RegExp_55.RegExp
    
    ' Création de l'objet d'Expression Régulière
    ' Le pattern permet de trouver tous les groupes de caractères suivants :
    ' (
    '       [^,\x22\r\n]*       => sans ',' '"' '\r' '\n' ou vide
    '   ou  \x22[^\x22]*\x22    => chaîne (pouvant être vide) sans '"' entourée de '"'
    '  )
    '  suivi de
    ' (
    '       ,     => virgule
    '   ou  \n    => retour à la ligne sous Linux
    '   ou  \r\n  => retour à la ligne sous Windows
    '   ou  \r    => ...je sais pas d'où il sort celui là, mais il n'est pas géré par Excel,
    '                donc le splitting de fichier ne fonctionne pas : il vaut mieux avoir un
    '                fichier ne dépassant pas 25/30 000 lignes
    '   ou  $     => fin de ligne ou fichier
    ' )
    reg.Pattern = "([^,\x22\r\n]*|\x22[^\x22]*\x22)(,|\n|\r\n|\r|$)"
    reg.Global = True

    Set GetCsvChunks = reg.Execute(file_content)
End Function

Function InjectCsvChunksIntoWorksheet(ByRef matches As VBScript_RegExp_55.MatchCollection, ByRef ws As Worksheet)
    Dim match As VBScript_RegExp_55.match
    Dim row_nb As Integer
    Dim col_letter As String
    Dim cell_content As String
    Dim separator As String
    
    row_nb = 1
    col_letter = "A"
    
    ' Parcours de chacun des blocs découpés précédemment
    For Each match In matches
    
        ' La première sous-partie du bloc sera le contenu de la cellule
        cell_content = TrimQuotes(match.SubMatches(0))
        ws.Range(col_letter & row_nb).Value = cell_content
        
        ' Incrémentation de la lettre de colonne
        col_letter = Chr(Asc(col_letter) + 1)
        
        ' Comme on lit le fichier en entier à cause du mode UTF-8 (et non pas ligne par ligne),
        ' on repère les changements de ligne lorsque le 2ème sous-bloc n'est pas une virgule
        separator = match.SubMatches(1)
        If separator <> "," Then
            row_nb = row_nb + 1
            col_letter = "A"
        End If
        
    Next match
End Function

Function TrimQuotes(ByVal str As String) As String
    Dim regQuotes As New VBScript_RegExp_55.RegExp
    regQuotes.Pattern = "\"""
    regQuotes.Global = True
    TrimQuotes = regQuotes.Replace(str, "")
End Function
