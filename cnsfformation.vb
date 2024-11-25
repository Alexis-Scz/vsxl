Sub Lancer_recuperation()

    'Dim wsOutput As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim FolderPath As String
    Dim Filename As String
    Dim i As Long
    Dim uniqueID  As String
    Dim serviceName As String
    Dim wsService  As Worksheet


    FolderPath = ThisWorkbook.Sheets(1).Range("C3").Value
    Filename = Dir(FolderPath & "\*.xls*")

    Do While Filename <> ""
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(FolderPath & "\" & Filename, UpdateLinks:=0)
        Application.DisplayAlerts = True
        Set ws = wb.Sheets(1) 'Assumer que les données sont dans la première feuille
        
        serviceName = ws.Range("C9").Value
        If serviceName = "" Then serviceName = "Rien"
            
        ' Vérifier si l'onglet pour ce service existe déjà
        On Error Resume Next
        Set wsService = ThisWorkbook.Sheets(serviceName)
        On Error GoTo 0

        ' Si l'onglet n'existe pas, en créer un
        If wsService Is Nothing Then
            Set wsService = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsService.Name = serviceName
            ' Mettre les en-têtes dans le nouvel onglet
            wsService.Cells(1, 1).Resize(1, 16).Value = Array("Unique id", _
            "intitulé 1","objectif 1", "avis 1","intitulé 2", "objectif 2",  "avis 2","intitulé 3", "objectif 3",  "avis 3",  "intitulé 4","objectif 4", "avis 4", _
           "intitulé 5","objectif 5",  "avis 5")

        End If

        ' Construction du uniqueID :
        uniqueID=serviceName
        uniqueID = uniqueID &" " &UCase(ws.Range("C8").Value)
        uniqueID = uniqueID & " " & ws.Range("M8").Value
        uniqueID = uniqueID & " " & Format(ws.Range("I6").Value, "dd-mm-yyyy")
        



        ' Trouver la dernière ligne vide dans l'onglet de service
        LastRow = wsService.Cells(wsService.Rows.Count, 1).End(xlUp).Row + 1
          
        trouve = 0
        For ligne = 1 To LastRow
            If wsService.Cells(ligne, 1).Value = uniqueID Then
                trouve = 1
                Exit For
            End If
        Next ligne
        
        ' Si pas trouvé alors on écrit à la fin de l'onglet
        If trouve = 0 Then
            With wsService
                .Cells(LastRow, 1).Value = uniqueID                ' UniqueID
                .Cells(LastRow, 3).Value = ws.Range("G130").Value  ' "intitulé 1"
                .Cells(LastRow, 2).Value = ws.Range("A130").Value  ' "objectif 1"
                .Cells(LastRow, 4).Value = ws.Range("N130").Value  ' "avis 1"
                .Cells(LastRow, 6).Value = ws.Range("G131").Value  ' "intitulé 2"
                .Cells(LastRow, 5).Value = ws.Range("A131").Value  ' "objectif 2"
                .Cells(LastRow, 7).Value = ws.Range("N131").Value  ' "avis 2"
                .Cells(LastRow, 9).Value = ws.Range("G132").Value  ' "intitulé 3"
                .Cells(LastRow, 8).Value = ws.Range("A132").Value  ' "objectif 3"
                .Cells(LastRow, 10).Value = ws.Range("N132").Value  ' "avis 3"
                .Cells(LastRow, 12).Value = ws.Range("G133").Value  ' "intitulé 4"
                .Cells(LastRow, 11).Value = ws.Range("A133").Value  ' "objectif 4"
                .Cells(LastRow, 13).Value = ws.Range("N133").Value  ' "avis 4"
                .Cells(LastRow, 15).Value = ws.Range("G134").Value  ' "intitulé 5"
                .Cells(LastRow, 14).Value = ws.Range("A134").Value  ' "objectif 5"
                .Cells(LastRow, 16).Value = ws.Range("N134").Value  ' "avis 5"



                ' Trouver la dernière colonne avec des données
                LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
                ' Définir la plage basée sur les dernières lignes et colonnes avec des données
                Set Rng = .Range("A1", .Cells(LastRow, LastCol))
                ' Ajuster la largeur des colonnes en fonction du contenu
                Rng.EntireColumn.AutoFit
                ' Ajouter des filtres sur les en-têtes
                Rng.AutoFilter
            End With
        End If
        
        ' Remettre à zéro la référence à la feuille de service pour le prochain tour de boucle
        Set wsService = Nothing
        
        
        wb.Close SaveChanges:=False
        'Renommez le fichier pour ajouter l'identifiant unique
       

        Filename = Dir
  Loop
stp:
   MsgBox "allé c bon!", vbInformation
        
End Sub

