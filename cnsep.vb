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
    Dim ent6ans     As String
    Dim formation6a     As String
    Dim certifpro6a     As String
    Dim progressionsala     As String

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
            wsService.Cells(1, 1).Resize(1, 74).Value = Array("Unique id", _
            "Date d'entretien","Date du dernier entretien","Nom","Prénom","Service","Fonction","Date d'embauche","Nom manager","Prénom manager","Fonction manager","3 entretien pro dans les 6 ans", _ 
            "date entretien 1","date entretien 2","date entretien 3","une formation dans les 6 ans","intitulé formation 1","obligatoire 1","date debut 1","date de fin 1", _
            "intitulé formation 2","obligatoire 2","date debut 2","date de fin 2","intitulé formation 3","obligatoire 3","date debut 3","date de fin 3", _
            "intitulé formation 4","obligatoire 4","date debut 4","date de fin 4","intitulé formation 5","obligatoire 5","date debut 5","date de fin 5", _
            "intitulé formation 6","obligatoire 6","date debut 6","date de fin 6","certif pro dans les 6ans","element de certif 1","intitulé 1","niveau 1","date obtention 1","dispositif mobilisé 1", _
            "element de certif 2","intitulé 2","niveau 2","date obtention 2","dispositif mobilisé 2","element de certif 3","intitulé 3","niveau 3","date obtention 3","dispositif mobilisé 3", _
            "element de certif 4","intitulé 4","niveau 4","date obtention 4","dispositif mobilisé 4","element de certif 5","intitulé 5","niveau 5","date obtention 5","dispositif mobilisé 5", _
            "element de certif 6","intitulé 6","niveau 6","date obtention 6","dispositif mobilisé 6","progress salariale ou pro ","date augmentation","%augmentation")

        End If
            'oui ou non

            ent6ans="pas repondu"
            formation6a="pas repondu"
            certifpro6a="pas repondu"
            progressionsala="pas repondu"

            if ws.range("B19").Value=true then
            ent6ans="oui"
            end if

            if ws.range("L19").value=true then
            ent6ans="non"
            end if





            if ws.range("B24").Value=true then
            formation6a="oui"
            end if

            if ws.range("L24").value=true then
            formation6a="non"
            end if





            if ws.range("B41").Value=true then
            certifpro6a="oui"
            end if

            if ws.range("L41").value=true then
            certifpro6a="non"
            end if





            if ws.range("B53").Value=true then
            progressionsala="oui"
            end if

            if ws.range("L53").value=true then
            progressionsala="non"
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
                .Cells(LastRow, 1).Value = uniqueID                      ' UniqueID
                .Cells(LastRow, 2).Value = ws.Range("J7").Value          ' "Date entretien actuel"
                .Cells(LastRow, 3).Value = ws.Range("O7").Value            '"Date du dernier entretien"
                .Cells(LastRow, 4).Value = ws.Range("D9").Value           ' '"Nom"
                .Cells(LastRow, 5).Value = ws.Range("N9").Value           ' '"Prénom"
                .Cells(LastRow, 6).Value = ws.Range("D10").Value           ' '"Service"
                .Cells(LastRow, 7).Value = ws.Range("G10").Value           ' '"Fonction"
                .Cells(LastRow, 8).Value = ws.Range("O10").Value           ' '"Date d'embauche"
                .Cells(LastRow, 9).Value = ws.Range("D12").Value           ' '"Nom manager"
                .Cells(LastRow, 10).Value = ws.Range("I12").Value           ' '"Prénom manager"
                .Cells(LastRow, 11).Value = ws.Range("O12").Value           ' '"Fonction manager"
                .Cells(LastRow, 12).Value = ent6ans                         ' '"3 entretien pro dans les 6 ans"
                .Cells(LastRow, 13).Value = ws.Range("E22").Value           '"date entretien 1"
                .Cells(LastRow, 14).Value = ws.Range("I22").Value           ''"date entretien 2"
                .Cells(LastRow, 15).Value = ws.Range("O22").Value           ''"date entretien 3"
                .Cells(LastRow, 16).Value = formation6a                      ''"une formation dans les 6 ans"
                .Cells(LastRow, 17).Value = ws.Range("B29").Value            ''"intitulé formation 1"
                .Cells(LastRow, 18).Value = ws.Range("G29").Value            ''"obligatoire 1"
                .Cells(LastRow, 19).Value = ws.Range("L29").Value            ''"date debut 1"
                .Cells(LastRow, 20).Value = ws.Range("P29").Value            ''"date de fin 1"
                .Cells(LastRow, 21).Value = ws.Range("B30").Value            '"intitulé formation 2"
                .Cells(LastRow, 22).Value = ws.Range("G30").Value            ''"obligatoire 2"
                .Cells(LastRow, 23).Value = ws.Range("L30").Value            ''"date debut 2"
                .Cells(LastRow, 24).Value = ws.Range("P30").Value            ''"date de fin 2"
                .Cells(LastRow, 25).Value = ws.Range("B31").Value            ''"intitulé formation 3"
                .Cells(LastRow, 26).Value = ws.Range("G31").Value            ''"obligatoire 3"
                .Cells(LastRow, 27).Value = ws.Range("L31").Value            ''"date debut 3"
                .Cells(LastRow, 28).Value = ws.Range("P31").Value            ''"date de fin 3"
                .Cells(LastRow, 29).Value = ws.Range("B32").Value            '"intitulé formation 4"
                .Cells(LastRow, 30).Value = ws.Range("G32").Value            ''"obligatoire 4"
                .Cells(LastRow, 31).Value = ws.Range("L32").Value            ''"date debut 4"
                .Cells(LastRow, 32).Value = ws.Range("P32").Value            ''"date de fin 4"
                .Cells(LastRow, 33).Value = ws.Range("B33").Value            ''"intitulé formation 5"
                .Cells(LastRow, 34).Value = ws.Range("G33").Value            ''"obligatoire 5"
                .Cells(LastRow, 35).Value = ws.Range("L33").Value            ''"date debut 5"
                .Cells(LastRow, 36).Value = ws.Range("P33").Value            ''"date de fin 5"
                .Cells(LastRow, 37).Value = ws.Range("B34").Value            '"intitulé formation 6"
                .Cells(LastRow, 38).Value = ws.Range("G34").Value            ''"obligatoire 6"
                .Cells(LastRow, 39).Value = ws.Range("L34").Value            ''"date debut 6"
                .Cells(LastRow, 40).Value = ws.Range("P34").Value            ''"date de fin 6"
                .Cells(LastRow, 41).Value = certifpro6a                          ''"certif pro dans les 6ans"
                .Cells(LastRow, 42).Value = ws.Range("B45").Value            ''"element de certif 1"
                .Cells(LastRow, 43).Value = ws.Range("E45").Value            ''"intitulé 1"
                .Cells(LastRow, 44).Value = ws.Range("H45").Value            ''"niveau 1"
                .Cells(LastRow, 45).Value = ws.Range("K45").Value            ''"date obtention 1"
                .Cells(LastRow, 46).Value = ws.Range("N45").Value            ''"dispositif mobilisé 1"
                .Cells(LastRow, 47).Value = ws.Range("B46").Value            '"element de certif 2"
                .Cells(LastRow, 48).Value = ws.Range("E46").Value            ''"intitulé 2"
                .Cells(LastRow, 49).Value = ws.Range("H46").Value            ''"niveau 2"
                .Cells(LastRow, 50).Value = ws.Range("K46").Value            ''"date obtention 2"
                .Cells(LastRow, 51).Value = ws.Range("N46").Value            ''"dispositif mobilisé 2"
                .Cells(LastRow, 52).Value = ws.Range("B47").Value            ''"element de certif 3"
                .Cells(LastRow, 53).Value = ws.Range("E47").Value            ''"intitulé 3"
                .Cells(LastRow, 54).Value = ws.Range("H47").Value            ''"niveau 3"
                .Cells(LastRow, 55).Value = ws.Range("K47").Value            ''"date obtention 3"
                .Cells(LastRow, 56).Value = ws.Range("N47").Value            ''"dispositif mobilisé 3"
                .Cells(LastRow, 57).Value = ws.Range("B48").Value            '"element de certif 4"
                .Cells(LastRow, 58).Value = ws.Range("E48").Value            ''"intitulé 4"
                .Cells(LastRow, 59).Value = ws.Range("H48").Value            ''"niveau 4"
                .Cells(LastRow, 60).Value = ws.Range("K48").Value            ''"date obtention 4"
                .Cells(LastRow, 61).Value = ws.Range("N48").Value            ''"dispositif mobilisé 4"
                .Cells(LastRow, 62).Value = ws.Range("B49").Value              '"element de certif 5"
                .Cells(LastRow, 63).Value = ws.Range("E49").Value            ''"intitulé 5"
                .Cells(LastRow, 64).Value = ws.Range("H49").Value            ''"niveau 5"
                .Cells(LastRow, 65).Value = ws.Range("K49").Value            ''"date obtention 5"
                .Cells(LastRow, 66).Value = ws.Range("N49").Value            ''"dispositif mobilisé 5"
                .Cells(LastRow, 67).Value = ws.Range("B50").Value            '"element de certif 6"
                .Cells(LastRow, 68).Value = ws.Range("E50").Value            ''"intitulé 6"
                .Cells(LastRow, 69).Value = ws.Range("H50").Value            ''"niveau 6"
                .Cells(LastRow, 70).Value = ws.Range("K50").Value            ''"date obtention 6"
                .Cells(LastRow, 71).Value = ws.Range("N50").Value            ''"dispositif mobilisé 6"
                .Cells(LastRow, 72).Value = progressionsala            ''"progress salariale ou pro "
                .Cells(LastRow, 73).Value = ws.Range("G56").Value            ''"date augmentation"
                .Cells(LastRow, 74).Value = ws.Range("O56").Value            ''"%augmentation"


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
        Name FolderPath & "\" & FileName As FolderPath & "\" & serviceName & " BILAN PARCOURS PRO 24-25 " & " " & uniqueID & ".xlsx"
        
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


