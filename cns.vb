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
    dim notem1     as string
    dim notem1     as string
    dim notem2     as string
    dim notem3     as string
    dim notem4     as string
    dim notem5     as string
    dim notem6     as string
    dim notem7     as string
    dim notem8     as string
    dim notecomp1  as string
    dim notecomp2  as string
    dim notecomp3  as string
    dim notecomp4  as string
    dim notecomp5  as string
    dim notecomp6  as string
    dim notecomp7  as string
    dim notecomp8  as string
    dim notecomp9  as string
    dim notecomp10 as string
    dim notecomp11 as string
    dim notecomp12 as string
    dim notecomp13 as string
    dim perfglo as string

    FolderPath = ThisWorkbook.sheets(1).Range("C3").value
    FileName = Dir(FolderPath & "\*.xls*")

    Do While FileName <> ""
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(FolderPath & "\" & FileName, UpdateLinks:=0)
        Application.DisplayAlerts = True
        Set ws = wb.Sheets(1) 'Assumer que les données sont dans la première feuille
        
        serviceName =ws.range("C9").Value 
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
            wsService.Cells(1, 1).Resize(1, 184).Value = Array( "Unique id", _
            "Date entretien actuel","date du dernier entretien","nom","prenon","service","fonction","date embauche","nom superieur","prenom superieur","fontion superieur", _
            "mission 1","notation mission 1","mission 2","notation mission 2","mission 3","notation mission 3","mission 4","notation mission 4","mission 5","notation mission 5", _
            "mission 6","notation mission 6","mission 7","notation mission 7","mission 8","notation mission 8","competences 1","notation competences 1","competences 2", _
            "notation competences 2","competences 3","notation competences 3","competences 4","notation competences 4","competences 5","notation competences 5","competences 6", _
            "notation competences 6","competences 7","notation competences 7","competences 8","notation competences 8","competences 9","notation competences 9","competences 10", _
            "notation competences 10","competences 11","notation competences 11","competences 12","notation competences 12","competences 13","notation competences 13","objectif acuel 1","realisation obj 1","commentaire 1","objectif actuel 2", _
            "realisation obj 2","commentaire 2","objectif actuel 3","realisation obj 3","commentaire 3","objectif actuel 4","realisation obj 4","commentaire 4","bilan global perf","new obj 1","new obj 2","new obj 3","new obj 4","comm osef", _
            "appreciation mana","commntaires salarie", _
            "diplome 1","diplome 2","diplome 3","diplome 4","diplome 5","langue 1","niveau 1","langue 2","niveau 2","langue 3","niveau 3","ancien poste 1","ancien activité 1","ancien date debut 1","ancien date fin 1", _
            "ancien poste 2","ancien activité 2","ancien date debut 2","ancien date fin 2","ancien poste 3","ancien activité 3","ancien date debut 3","ancien date fin 3","ancien poste 4","ancien activité 4","ancien date debut 4","ancien date fin 4", _
            "ancien poste 5","ancien activité 5","ancien date debut 5","ancien date fin 5","poste 1","activité 1","date debut 1","date fin 1", _
            "poste 2","activité 2","date debut 2","date fin 2","poste 3","activité 3","date debut 3","date fin 3","poste 4","activité 4","date debut 4","date fin 4", _
            "poste 5","activité 5","date debut 5","date fin 5", _
            "formation 1","date de debut 1","date de fin 1","formation 2","date de debut 2","date de fin 2","formation 3","date de debut 3","date de fin 3", _
            "formation 4","date de debut 4","date de fin 4","formation 5","date de debut 5","date de fin 5","apport 1","apport 2","apport 3","apport 4", _
           "Bilan de compétences Date debut","Bilan de compétences Date fin","Bilan de compétences comm salarie","Entretien avec un Conseiller en évolution Professionnel? Date debut", _
           "Entretien avec un Conseiller en évolution Professionnel? Date fin","Entretien avec un Conseiller en évolution Professionnel? comm salarie", _
           "Compte Personnel de Formation? CPF Date debut","Compte Personnel de Formation? CPF Date fin","Compte Personnel de Formation? CPF comm salarie", _
           "Validation des acquis et de l'expérience Date debut","Validation des acquis et de l'expérience Date fin","Validation des acquis et de l'expérience comm salarie", _
           "Autres - précisez Date debut","Autres - précisez Date fin","Autres - précisez comm salarie","souhait 1","avis responsable 1","souhait 2","avis responsable 2","souhait 3","avis responsable 3", _
           "souhait 4","avis responsable 4","objectif 1","intitulé 1","avis 1","objectif 2","intitulé 2","avis 2","objectif 3","intitulé 3","avis 3","objectif 4","intitulé 4","avis 4", _
           "objectif 5","intitulé 5","avis 5","commentaire collabo","commentaire responsable")

        End If

        ' Construction du uniqueID :
        uniqueID = UCase(ws.Range("C8").Value)
        uniqueID = uniqueID & " " & ws.Range("M8").Value
        uniqueID = uniqueID & " " & Format(ws.Range("I6").Value,"dd-mm-yyyy")
        

        notem1    ="Pas noté"
        notem1    ="Pas noté"
        notem2    ="Pas noté"
        notem3    ="Pas noté"
        notem4    ="Pas noté"
        notem5    ="Pas noté"
        notem6    ="Pas noté"
        notem7    ="Pas noté"
        notem8    ="Pas noté"
        notecomp1 ="Pas noté"
        notecomp2 ="Pas noté"
        notecomp3 ="Pas noté"
        notecomp4 ="Pas noté"
        notecomp5 ="Pas noté"
        notecomp6 ="Pas noté"
        notecomp7 ="Pas noté"
        notecomp8 ="Pas noté"
        notecomp9 ="Pas noté"
        notecomp10="Pas noté"
        notecomp11="Pas noté"
        notecomp12="Pas noté"
        notecomp13="Pas noté"
        perfglo="pas noté"

        'notes
        if ws.range("G16").Value=true then
            notem1="excellent"
        end if

        if ws.range("H16").value=true then
            notem1="Bien"
        end if

        if ws.range("I16").value=true Then
            notem1="Moyen"
        End if 

        if ws.range("J16").value=true then
            notem1="Insuffisant"
        End if
        



        if ws.range("G17").Value=true then
            notem2="excellent"
        end if

        if ws.range("H17").value=true then
            notem2="Bien"
        end if

        if ws.range("I17").value=true Then
            notem2="Moyen"
        End if 

        if ws.range("J17").value=true then
            notem2="Insuffisant"
        End if




        if ws.range("G18").Value=true then
            notem3="excellent"
        end if

        if ws.range("H18").value=true then
            notem3="Bien"
        end if

        if ws.range("I18").value=true Then
            notem3="Moyen"
        End if 

        if ws.range("J18").value=true then
            notem3="Insuffisant"
        End if




        if ws.range("G19").Value=true then
            notem4="excellent"
        end if

        if ws.range("H19").value=true then
            notem4="Bien"
        end if

        if ws.range("I19").value=true Then
            notem4="Moyen"
        End if 

        if ws.range("J19").value=true then
            notem4="Insuffisant"
        End if




        if ws.range("G20").Value=true then
            notem5="excellent"
        end if

        if ws.range("H20").value=true then
            notem5="Bien"
        end if

        if ws.range("I20").value=true Then
            notem5="Moyen"
        End if 

        if ws.range("J20").value=true then
            notem5="Insuffisant"
        End if




        if ws.range("G21").Value=true then
            notem6="excellent"
        end if

        if ws.range("H21").value=true then
            notem6="Bien"
        end if

        if ws.range("I21").value=true Then
            notem6="Moyen"
        End if 

        if ws.range("J21").value=true then
            notem6="Insuffisant"
        End if




        if ws.range("G22").Value=true then
            notem7="excellent"
        end if

        if ws.range("H22").value=true then
            notem7="Bien"
        end if

        if ws.range("I22").value=true Then
            notem7="Moyen"
        End if 

        if ws.range("J22").value=true then
            notem7="Insuffisant"
        End if






        if ws.range("G23").Value=true then
            notem8="excellent"
        end if

        if ws.range("H23").value=true then
            notem8="Bien"
        end if

        if ws.range("I23").value=true Then
            notem8="Moyen"
        End if 

        if ws.range("J23").value=true then
            notem8="Insuffisant"
        End if






        if ws.range("G24").Value=true then
            notem9="excellent"
        end if

        if ws.range("H24").value=true then
            notem9="Bien"
        end if

        if ws.range("I24").value=true Then
            notem9="Moyen"
        End if 

        if ws.range("J24").value=true then
            notem9="Insuffisant"
        End if





        if ws.range("G26").Value=true then
            notecomp1="excellent"
        end if

        if ws.range("H26").value=true then
            notecomp1="Bien"
        end if

        if ws.range("I26").value=true Then
            notecomp1="Moyen"
        End if 

        if ws.range("J26").value=true then
            notecomp1="Insuffisant"
        End if



        if ws.range("G27").Value=true then
            notecomp2="excellent"
        end if

        if ws.range("H27").value=true then
            notecomp2="Bien"
        end if

        if ws.range("I27").value=true Then
            notecomp2="Moyen"
        End if 

        if ws.range("J27").value=true then
            notecomp2="Insuffisant"
        End if



        if ws.range("G28").Value=true then
            notecomp3="excellent"
        end if

        if ws.range("H28").value=true then
            notecomp3="Bien"
        end if

        if ws.range("I28").value=true Then
            notecomp3="Moyen"
        End if 

        if ws.range("J28").value=true then
            notecomp3="Insuffisant"
        End if



        if ws.range("G29").Value=true then
            notecomp4="excellent"
        end if

        if ws.range("H29").value=true then
            notecomp4="Bien"
        end if

        if ws.range("I29").value=true Then
            notecomp4="Moyen"
        End if 

        if ws.range("J29").value=true then
            notecomp4="Insuffisant"
        End if



        if ws.range("G30").Value=true then
            notecomp5="excellent"
        end if

        if ws.range("H30").value=true then
            notecomp5="Bien"
        end if

        if ws.range("I30").value=true Then
            notecomp5="Moyen"
        End if 

        if ws.range("J30").value=true then
            notecomp5="Insuffisant"
        End if





        if ws.range("G31").Value=true then
            notecomp6="excellent"
        end if

        if ws.range("H31").value=true then
            notecomp6="Bien"
        end if

        if ws.range("I31").value=true Then
            notecomp6="Moyen"
        End if 

        if ws.range("J31").value=true then
            notecomp6="Insuffisant"
        End if



        if ws.range("G32").Value=true then
            notecomp7="excellent"
        end if

        if ws.range("H32").value=true then
            notecomp7="Bien"
        end if

        if ws.range("I32").value=true Then
            notecomp7="Moyen"
        End if 

        if ws.range("J32").value=true then
            notecomp7="Insuffisant"
        End if



        if ws.range("G33").Value=true then
            notecomp8="excellent"
        end if

        if ws.range("H33").value=true then
            notecomp8="Bien"
        end if

        if ws.range("I33").value=true Then
            notecomp8="Moyen"
        End if 

        if ws.range("J33").value=true then
            notecomp8="Insuffisant"
        End if



        if ws.range("G34").Value=true then
            notecomp9="excellent"
        end if

        if ws.range("H34").value=true then
            notecomp9="Bien"
        end if

        if ws.range("I34").value=true Then
            notecomp9="Moyen"
        End if 

        if ws.range("J34").value=true then
            notecomp9="Insuffisant"
        End if



        if ws.range("G35").Value=true then
            notecomp10="excellent"
        end if

        if ws.range("H35").value=true then
            notecomp10="Bien"
        end if

        if ws.range("I35").value=true Then
            notecomp10="Moyen"
        End if 

        if ws.range("J35").value=true then
            notecomp10="Insuffisant"
        End if



        if ws.range("G36").Value=true then
            notecomp11="excellent"
        end if

        if ws.range("H36").value=true then
            notecomp11="Bien"
        end if

        if ws.range("I36").value=true Then
            notecomp11="Moyen"
        End if 

        if ws.range("J36").value=true then
            notecomp11="Insuffisant"
        End if



        if ws.range("G37").Value=true then
            notecomp12="excellent"
        end if

        if ws.range("H37").value=true then
            notecomp12="Bien"
        end if

        if ws.range("I37").value=true Then
            notecomp12="Moyen"
        End if 

        if ws.range("J37").value=true then
            notecomp12="Insuffisant"
        End if



        if ws.range("G38").Value=true then
            notecomp13="excellent"
        end if

        if ws.range("H38").value=true then
            notecomp13="Bien"
        end if

        if ws.range("I38").value=true Then
            notecomp13="Moyen"
        End if 

        if ws.range("J38").value=true then
            notecomp13="Insuffisant"
        End if



        if ws.range("A47").Value=true then
            perfglo="excellent"
        end if

        if ws.range("F47").value=true then
            perfglo="Bien"
        end if

        if ws.range("K47").value=true Then
            perfglo="Moyen"
        End if 

        if ws.range("P47").value=true then
            perfglo="Insuffisant"
        End if


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
                .Cells(LastRow, 2).Value = ws.range("I6").Value  ' "Date entretien actuel"
                .Cells(LastRow, 3).Value = ws.range("Q6").Value  ' "date du dernier entretien"
                .Cells(LastRow, 4).Value = ws.range("C8").Value  ' "nom"
                .Cells(LastRow, 5).Value = ws.range("M8").Value  ' "prenon"
                .Cells(LastRow, 6).Value = ws.range("C9").Value  ' "service"
                .Cells(LastRow, 7).Value = ws.range("H9").Value  ' "fonction"
                .Cells(LastRow, 8).Value = ws.range("R9").Value  ' "date embauche"
                .Cells(LastRow, 9).Value = ws.range("C11").Value  ' "nom superieur"
                .Cells(LastRow, 10).Value = ws.range("I11").Value  ' "prenom superieur"
                .Cells(LastRow, 11).Value = ws.range("P11").Value  ' "fontion superieur"
                .Cells(LastRow, 12).Value = ws.range("A16").Value  ' "mission 1"
                .Cells(LastRow, 13).Value =notem1                                  ' "notation mission 1"
                .Cells(LastRow, 14).Value = ws.range("A17").Value  ' "mission 2"
                .Cells(LastRow, 15).Value =notem2                                    ' "notation mission 2"
                .Cells(LastRow, 16).Value = ws.range("A18").Value  ' "mission 3"
                .Cells(LastRow, 17).Value =notem3                                  ' "notation mission 3"
                .Cells(LastRow, 18).Value = ws.range("A19").Value  ' "mission 4"
                .Cells(LastRow, 19).Value = notem4                                 ' "notation mission 4"
                .Cells(LastRow, 20).Value = ws.range("A20").Value  ' "mission 5"
                .Cells(LastRow, 21).Value =   notem5                               ' "notation mission 2"
                .Cells(LastRow, 22).Value = ws.range("A21").Value  ' "mission 6"
                .Cells(LastRow, 23).Value =  notem6                                ' "notation mission 2"
                .Cells(LastRow, 24).Value = ws.range("A22").Value  ' "mission 7"
                .Cells(LastRow, 25).Value =  notem7                                ' "notation mission 2 "
                .Cells(LastRow, 26).Value = ws.range("A23").Value  ' "mission 8"
                .Cells(LastRow, 27).Value =notem8                          ' "notation mission 2"
                .Cells(LastRow, 28).Value = ws.range("A26").Value  ' "competences 1"
                .Cells(LastRow, 29).Value =notecomp1                                  ' "notation competences 1"
                .Cells(LastRow, 30).Value = ws.range("A27").Value  ' "competences 2"
                .Cells(LastRow, 31).Value =notecomp2                                  ' "notation competences 1"
                .Cells(LastRow, 32).Value = ws.range("A28").Value  ' "competences 3"
                .Cells(LastRow, 33).Value =notecomp3                                  ' "notation competences 1"
                .Cells(LastRow, 34).Value = ws.range("A29").Value  ' "competences 4"
                .Cells(LastRow, 35).Value =notecomp4                                  ' "notation competences 1"
                .Cells(LastRow, 36).Value = ws.range("A30").Value  ' "competences 5"
                .Cells(LastRow, 37).Value = notecomp5                                 ' "notation competences 1"
                .Cells(LastRow, 38).Value = ws.range("A31").Value  ' "competences 6"
                .Cells(LastRow, 39).Value =notecomp6                                  ' "notation competences 1"
                .Cells(LastRow, 40).Value = ws.range("A32").Value  ' "competences 7"
                .Cells(LastRow, 41).Value = notecomp7                                 ' "notation competences 1"
                .Cells(LastRow, 42).Value = ws.range("A33").Value  ' "competences 8"
                .Cells(LastRow, 43).Value =notecomp8                                  ' "notation competences 1"
                .Cells(LastRow, 44).Value = ws.range("A34").Value  ' "competences 9"
                .Cells(LastRow, 45).Value = notecomp9                                 ' "notation competences 1"
                .Cells(LastRow, 46).Value = ws.range("A35").Value  ' "competences 10"
                .Cells(LastRow, 47).Value = notecomp10                                 ' "notation competences 1"
                .Cells(LastRow, 48).Value = ws.range("A36").Value  ' "competences 11"
                .Cells(LastRow, 49).Value =  notecomp11                                ' "notation competences 1"
                .Cells(LastRow, 50).Value = ws.range("A37").Value  ' "competences 12"
                .Cells(LastRow, 51).Value = notecomp12                                 ' "notation competences 12"
                .Cells(LastRow, 52).Value = ws.range("A38").Value  ' "competences 13"
                .Cells(LastRow, 53).Value = notecomp13                                 ' "notation competences 12"
                .Cells(LastRow, 54).Value = ws.range("A41").Value  ' "objectif acuel 1"
                .Cells(LastRow, 55).Value = ws.range("G41").Value  ' "realisation obj 1"
                .Cells(LastRow, 56).Value = ws.range("J41").Value  '"commentaire 1"  
                .Cells(LastRow, 57).Value = ws.range("A42").Value  ' "objectif actuel 2"
                .Cells(LastRow, 58).Value = ws.range("G42").Value  ' "realisation obj 2"
                .Cells(LastRow, 59).Value = ws.range("J42").Value  '"commentaire 2"
                .Cells(LastRow, 60).Value = ws.range("A43").Value  ' "objectif actuel 3"
                .Cells(LastRow, 61).Value = ws.range("G43").Value  ' "realisation obj 3"
                .Cells(LastRow, 62).Value = ws.range("J43").Value  '"commentaire 3"
                .Cells(LastRow, 63).Value = ws.range("A44").Value  ' "objectif actuel 4"
                .Cells(LastRow, 64).Value = ws.range("G44").Value  ' "realisation obj 4"
                .Cells(LastRow, 65).Value = ws.range("J44").Value  ' "commentaire 4"
                .Cells(LastRow, 66).Value = perfglo                ' "bilan global perf"
                .Cells(LastRow, 67).Value = ws.range("B49").Value  ' "new obj 1"
                .Cells(LastRow, 68).Value = ws.range("B50").Value  ' "new obj 2"
                .Cells(LastRow, 69).Value = ws.range("B51").Value  ' "new obj 3"
                .Cells(LastRow, 70).Value = ws.range("B52").Value  ' "new obj 4"
                .Cells(LastRow, 71).Value = ws.range("A54").Value  ' "comm osef"
                .Cells(LastRow, 72).Value = ws.range("A56").Value  ' "appreciation mana"
                .Cells(LastRow, 73).Value = ws.range("A58").Value  ' "commntaires salarie"
                .Cells(LastRow, 74).Value = ws.range("F74").Value  ' "diplome 1"
                .Cells(LastRow, 75).Value = ws.range("F75").Value  ' "diplome 2"
                .Cells(LastRow, 76).Value = ws.range("F76").Value  ' "diplome 3"
                .Cells(LastRow, 77).Value = ws.range("F77").Value  ' "diplome 4"
                .Cells(LastRow, 78).Value = ws.range("F78").Value  ' "diplome 5"
                .Cells(LastRow, 79).Value = ws.range("F80").Value  ' "langue 1"
                .Cells(LastRow, 80).Value = ws.range("N80").Value  ' "niveau 1"
                .Cells(LastRow, 81).Value = ws.range("F82").Value  ' "langue 2"
                .Cells(LastRow, 82).Value = ws.range("N82").Value  ' "niveau 2"
                .Cells(LastRow, 83).Value = ws.range("F84").Value  ' "langue 3"
                .Cells(LastRow, 84).Value = ws.range("N84").Value  ' "niveau 3"
                .Cells(LastRow, 85).Value = ws.range("A89").Value  ' "ancien poste 1"
                .Cells(LastRow, 86).Value = ws.range("F89").Value  ' "ancien activité 1"
                .Cells(LastRow, 87).Value = ws.range("N89").Value  ' "ancien date debut 1"
                .Cells(LastRow, 88).Value = ws.range("Q89").Value  ' "ancien date fin 1"
                .Cells(LastRow, 89).Value = ws.range("A90").Value  ' "ancien poste 2"
                .Cells(LastRow, 90).Value = ws.range("F90").Value  ' "ancien activité 2"
                .Cells(LastRow, 91).Value = ws.range("N90").Value  ' "ancien date debut 2"
                .Cells(LastRow, 92).Value = ws.range("Q90").Value  ' "ancien date fin 2"
                .Cells(LastRow, 93).Value = ws.range("A91").Value  ' "ancien poste 3"
                .Cells(LastRow, 94).Value = ws.range("F91").Value  ' "ancien activité 3"
                .Cells(LastRow, 95).Value = ws.range("N91").Value  ' "ancien date debut 3"
                .Cells(LastRow, 96).Value = ws.range("Q91").Value  ' "ancien date fin 3"
                .Cells(LastRow, 97).Value = ws.range("A92").Value  ' "ancien poste 4"
                .Cells(LastRow, 98).Value = ws.range("F92").Value  ' "ancien activité 4"
                .Cells(LastRow, 99).Value = ws.range("N92").Value  ' "ancien date debut 4"
                .Cells(LastRow, 100).Value = ws.range("Q92").Value  ' "ancien date fin 4"
                .Cells(LastRow, 101).Value = ws.range("A93").Value  ' "ancien poste 5"
                .Cells(LastRow, 102).Value = ws.range("F93").Value  ' "ancien activité 5"
                .Cells(LastRow, 103).Value = ws.range("N93").Value  ' "ancien date debut 5"
                .Cells(LastRow, 104).Value = ws.range("Q93").Value  ' "ancien date fin 5"
                .Cells(LastRow, 105).Value = ws.range("A96").Value  ' "poste 1"
                .Cells(LastRow, 106).Value = ws.range("F96").Value  ' "activité 1"
                .Cells(LastRow, 107).Value = ws.range("N96").Value  ' "date debut 1"
                .Cells(LastRow, 108).Value = ws.range("Q96").Value  ' "date fin 1"
                .Cells(LastRow, 109).Value = ws.range("A97").Value  ' "poste 2"
                .Cells(LastRow, 110).Value = ws.range("F97").Value  ' "activité 2"
                .Cells(LastRow, 111).Value = ws.range("N97").Value  ' "date debut 2"
                .Cells(LastRow, 112).Value = ws.range("Q97").Value  ' "date fin 2"
                .Cells(LastRow, 113).Value = ws.range("A98").Value  ' "poste 3"
                .Cells(LastRow, 114).Value = ws.range("F98").Value  ' "activité 3"
                .Cells(LastRow, 115).Value = ws.range("N98").Value  ' "date debut 3"
                .Cells(LastRow, 116).Value = ws.range("Q98").Value  ' "date fin 3"
                .Cells(LastRow, 117).Value = ws.range("A99").Value  ' "poste 4"
                .Cells(LastRow, 118).Value = ws.range("F99").Value  ' "activité 4"
                .Cells(LastRow, 119).Value = ws.range("N99").Value  ' "date debut 4"
                .Cells(LastRow, 120).Value = ws.range("Q99").Value  ' "date fin 4"
                .Cells(LastRow, 121).Value = ws.range("A100").Value  ' "poste 5"
                .Cells(LastRow, 122).Value = ws.range("F100").Value  ' "activité 5"
                .Cells(LastRow, 123).Value = ws.range("N100").Value  ' "date debut 5"
                .Cells(LastRow, 124).Value = ws.range("Q100").Value  ' "date fin 5"
                .Cells(LastRow, 125).Value = ws.range("A104").Value  ' "formation 1"
                .Cells(LastRow, 126).Value = ws.range("N104").Value  ' "date de debut 1"
                .Cells(LastRow, 127).Value = ws.range("Q104").Value  ' "date de fin 1"
                .Cells(LastRow, 128).Value = ws.range("A105").Value  ' "formation 2"
                .Cells(LastRow, 129).Value = ws.range("N105").Value  ' "date de debut 2"
                .Cells(LastRow, 130).Value = ws.range("Q105").Value  ' "date de fin 2"
                .Cells(LastRow, 131).Value = ws.range("A106").Value  ' "formation 3"
                .Cells(LastRow, 132).Value = ws.range("N106").Value  ' "date de debut 3"
                .Cells(LastRow, 133).Value = ws.range("Q106").Value  ' "date de fin 3"
                .Cells(LastRow, 134).Value = ws.range("A107").Value  ' "formation 4"
                .Cells(LastRow, 135).Value = ws.range("N107").Value  ' "date de debut 4"
                .Cells(LastRow, 136).Value = ws.range("Q107").Value  ' "date de fin 4"
                .Cells(LastRow, 137).Value = ws.range("A108").Value  ' "formation 5"
                .Cells(LastRow, 138).Value = ws.range("N108").Value  ' "date de debut 5"
                .Cells(LastRow, 139).Value = ws.range("Q108").Value  ' "date de fin 5"
                .Cells(LastRow, 140).Value = ws.range("A110").Value  ' "apport 1"
                .Cells(LastRow, 141).Value = ws.range("A111").Value  ' "apport 2"
                .Cells(LastRow, 142).Value = ws.range("A112").Value  ' "apport 3"
                .Cells(LastRow, 143).Value = ws.range("A113").Value  ' "apport 4"
                .Cells(LastRow, 144).Value = ws.range("F116").Value  ' "Bilan de compétences Date debut"
                .Cells(LastRow, 145).Value = ws.range("I116").Value  ' "Bilan de compétences Date fin"
                .Cells(LastRow, 146).Value = ws.range("L116").Value  ' "Bilan de compétences comm salarie"
                .Cells(LastRow, 147).Value = ws.range("F117").Value  ' "Entretien avec un Conseiller en évolution Professionnel? Date debut"
                .Cells(LastRow, 148).Value = ws.range("I117").Value  ' "Entretien avec un Conseiller en évolution Professionnel? Date fin"
                .Cells(LastRow, 149).Value = ws.range("L117").Value  ' "Entretien avec un Conseiller en évolution Professionnel? comm salarie"
                .Cells(LastRow, 150).Value = ws.range("F118").Value  ' "Compte Personnel de Formation? CPF Date debut"
                .Cells(LastRow, 151).Value = ws.range("I118").Value  ' "Compte Personnel de Formation? CPF Date fin"
                .Cells(LastRow, 152).Value = ws.range("L118").Value  ' "Compte Personnel de Formation? CPF comm salarie"
                .Cells(LastRow, 153).Value = ws.range("F119").Value  ' "Validation des acquis et de l'expérience Date debut"
                .Cells(LastRow, 154).Value = ws.range("I119").Value  ' "Validation des acquis et de l'expérience Date fin"
                .Cells(LastRow, 155).Value = ws.range("L119").Value  ' "Validation des acquis et de l'expérience comm salarie"
                .Cells(LastRow, 156).Value = ws.range("F120").Value  ' "Autres - précisez Date debut"
                .Cells(LastRow, 157).Value = ws.range("I120").Value  ' "Autres - précisez Date fin"
                .Cells(LastRow, 158).Value = ws.range("L120").Value  ' "Autres - précisez comm salarie"
                .Cells(LastRow, 159).Value = ws.range("A124").Value  ' "souhait 1"
                .Cells(LastRow, 160).Value = ws.range("J124").Value  ' "avis responsable 1"
                .Cells(LastRow, 161).Value = ws.range("A125").Value  ' "souhait 2"
                .Cells(LastRow, 162).Value = ws.range("J125").Value  ' "avis responsable 2"
                .Cells(LastRow, 163).Value = ws.range("A126").Value  ' "souhait 3"
                .Cells(LastRow, 164).Value = ws.range("J126").Value  ' "avis responsable 3"
                .Cells(LastRow, 165).Value = ws.range("A127").Value  ' "souhait 4"
                .Cells(LastRow, 166).Value = ws.range("J127").Value  ' "avis responsable 4"
                .Cells(LastRow, 167).Value = ws.range("A130").Value  ' "objectif 1"
                .Cells(LastRow, 168).Value = ws.range("G130").Value  ' "intitulé 1"
                .Cells(LastRow, 169).Value = ws.range("N130").Value  ' "avis 1"
                .Cells(LastRow, 170).Value = ws.range("A131").Value  ' "objectif 2"
                .Cells(LastRow, 171).Value = ws.range("G131").Value  ' "intitulé 2"
                .Cells(LastRow, 172).Value = ws.range("N131").Value  ' "avis 2"
                .Cells(LastRow, 173).Value = ws.range("A132").Value  ' "objectif 3"
                .Cells(LastRow, 174).Value = ws.range("G132").Value  ' "intitulé 3"
                .Cells(LastRow, 175).Value = ws.range("N132").Value  ' "avis 3"
                .Cells(LastRow, 176).Value = ws.range("A133").Value  ' "objectif 4"
                .Cells(LastRow, 177).Value = ws.range("G133").Value  ' "intitulé 4"
                .Cells(LastRow, 178).Value = ws.range("N133").Value  ' "avis 4"
                .Cells(LastRow, 179).Value = ws.range("A134").Value  ' "objectif 5"
                .Cells(LastRow, 180).Value = ws.range("G134").Value  ' "intitulé 5"
                .Cells(LastRow, 181).Value = ws.range("N134").Value  ' "avis 5"
                .Cells(LastRow, 182).Value = ws.range("A137").Value  ' "commentaire collabo"
                .Cells(LastRow, 183).Value = ws.range("J137").Value  ' "commentaire responsable"


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
        Name FolderPath & "\" & FileName As FolderPath & "\" & serviceName & " EA 24-25 " & " " & uniqueID & ".xlsx"
        
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