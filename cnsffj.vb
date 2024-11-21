Sub Lancer_recuperation()

    'Dim wsOutput As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim FolderPath As String
    Dim FileName As String
    Dim i As Long
    Dim uniqueID  As String
    Dim serviceName As String
    Dim wsService  As Worksheet
    Dim  eqproperso   As String
    Dim  chargeampli  As String
    Dim  droitdeco    As String
    Dim  reposmin     As String
    dim  remunok      As string
    FolderPath = ThisWorkbook.Sheets(1).Range("C3").Value
    FileName = Dir(FolderPath & "\*.xls*")

    Do While FileName <> ""
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(FolderPath & "\" & FileName, UpdateLinks:=0)
        Application.DisplayAlerts = True
        Set ws = wb.Sheets(1) 'Assumer que les données sont dans la première feuille
        
        serviceName = ws.Range("D10").Value
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
            wsService.Cells(1, 1).Resize(1, 21).Value = Array("Unique id", _
            "Date d'entretien","date dernier entretien","Nom","Prénom","Service","Fonction","Date d'embauche","nom manager","prenom manager","fonction manager","eqproperso", _
            "commentaire  eqproperso","chargeampli","commentaire chargeampli","droitdeco","commentaire droitdeco","reposmin","commentaire reposmin","remunok","commentaire remunok")

        End If


            'oui ou non

            eqproperso  ="pas repondu"
            chargeampli ="pas repondu"
            droitdeco   ="pas repondu"
            reposmin    ="pas repondu"
            remunok     ="pas repondu"

            if ws.range("J18").Value=true then
            eqproperso="oui"
            end if

            if ws.range("K18").value=true then
            eqproperso="non"
            end if

            if ws.range("J19").Value=true then
            chargeampli="oui"
            end if

            if ws.range("K19").value=true then
            chargeampli="non"
            end if

            if ws.range("J20").Value=true then
            droitdeco="oui"
            end if

            if ws.range("K20").value=true then
            droitdeco="non"
            end if

            if ws.range("J21").Value=true then
            reposmin="oui"
            end if

            if ws.range("K21").value=true then
            reposmin="non"
            end if

            if ws.range("J22").Value=true then
            remunok="oui"
            end if

            if ws.range("K22").value=true then
            remunok="non"
            end if


        ' Construction du uniqueID :
        uniqueID = UCase(ws.Range("D9").Value)
        uniqueID = uniqueID & " " & ws.Range("N9").Value
        uniqueID = uniqueID & " " & Format(ws.Range("J7").Value, "dd-mm-yyyy")
        

      

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
                .Cells(LastRow, 1).Value = uniqueID                                       ' UniqueID'"Unique id"
                .Cells(LastRow, 2).Value = ws.Range("J7").Value                           '"Date d'entretien"
                .Cells(LastRow, 3).Value = ws.Range("Q7").Value                           '"date dernier entretien"
                .Cells(LastRow, 4).Value = ws.Range("D9").Value                           '"Nom"
                .Cells(LastRow, 5).Value = ws.Range("N9").Value                           '"Prénom"
                .Cells(LastRow, 6).Value = ws.Range("D10").Value                          '"Service"
                .Cells(LastRow, 7).Value = ws.Range("I10").Value                          '"Fonction"
                .Cells(LastRow, 8).Value = ws.Range("R10").Value                          '"Date d'embauche"
                .Cells(LastRow, 9).Value = ws.Range("D12").Value                          '"nom manager"
                .Cells(LastRow, 10).Value = ws.Range("J12").Value                          '"prenom manager"
                .Cells(LastRow, 11).Value = ws.Range("Q12").Value                          '"fonction manager"
                .Cells(LastRow, 12).Value = eqproperso                          ' eqproperso 
                .Cells(LastRow, 13).Value=ws.range("L18").Value                 'commentaire  eqproperso 
                .Cells(LastRow, 14).Value = chargeampli                         ' chargeampli
                .Cells(LastRow, 15).Value=ws.range("L19").Value                 'commentaire chargeampli
                .Cells(LastRow, 16).Value = droitdeco                           ' droitdeco  
                .Cells(LastRow, 17).Value=ws.range("L20").Value                 'commentaire droitdeco  
                .Cells(LastRow, 18).Value = reposmin                            ' reposmin   
                .Cells(LastRow, 19).Value=ws.range("L21").Value                 'commentaire reposmin   
                .Cells(LastRow, 20).Value = remunok                             ' remunok    
                .Cells(LastRow, 21).Value=ws.range("L22").Value                 'commentaire remunok    

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
        Name FolderPath & "\" & FileName As FolderPath & "\" & serviceName & " FORFAIT JOUR 24-25 " & " " & uniqueID & ".xlsx"
        
        FileName = Dir
Loop



    MsgBox "Extraction terminée!", vbInformation

End Sub



Sub DeleteAllSheetsExceptMacro()
    
    If ThisWorkbook.Worksheets.Count > 6 Then
        
        msg = MsgBox("Etes-vous sûr ?? " & vbCrLf & "Les onglets seront supprimés ainsi que les lignes qu'ils contiennent (sauf onglet Macro et stats)." & vbCrLf & "   Il faudra tout récupérer de nouveau...", vbOKCancel, "On supprime tout ?")
        If msg = vbOK Then
            Dim ws As Worksheet
            Application.DisplayAlerts = False
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name <> "Macro" Then
                    If ws.Name <> "Stats" Then
                        If ws.Name <> "Listes" Then
                            If ws.Name <> "Calculs" Then
                                If ws.Name <> "Total" Then
                                    If ws.Name <> "Fonction " Then
                                        If ws.Name <> "Admin" Then
                                            ws.Delete
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next ws
            Application.DisplayAlerts = True
        End If
    End If
End Sub


