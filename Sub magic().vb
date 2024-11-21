Sub magic()


    Dim n As Integer
    Dim a As Integer
    Dim der As Integer
    Dim job As String
    Dim classif As String
    Dim b As Integer
    Dim dercls As Integer
    Dim cazclassf As Range
    Dim lgn As Integer
    Dim coln As Integer
    Dim cazjob As Range
    dim lgnjob as Integer
    dim coljob as Integer


    dercls = Sheets("DATA BASE").Range("J" & Rows.Count).End(xlUp).Row
    der = Sheets("DATA BASE").Range("A" & Rows.Count).End(xlUp).Row


    For n = 2 To dercls
        i = n
        classif = Sheets("DATA BASE").Range("J" & n)
        Sheets("MAPPING").Range("A" & n).Value = classif

    Next n

    b = 2
    For a = 2 To der

        job = Sheets("DATA BASE").Range("A" & a).Value
        classif = Sheets("DATA BASE").Range("B" & a)


        Set cazjob = Sheets("MAPPING").Range(Cells(dercls + 1, b), Cells(dercls + 1, b)).Find(job)
        If cazjob Is Nothing Then
            Sheets("MAPPING").Range(Cells(dercls + 1, b), Cells(dercls + 1, b)).Value = job
            lgnjob=dercls+1
            coljob=b
            b=b+1
        else
            lgnjob=cazjob.Row
            coljob=cazjob.Column


        End If
        Set cazclassf = Sheets("MAPPING").Range(Cells(2, 1), Cells(dercls, 2)).Find(classif)
        lgn = cazclassf.Row
        coln = cazclassf.Column
        Sheets("MAPPING").Range(Cells(lgn, coljob), Cells(lgn, coljob)).Value = "X"



    Next a


End Sub

sub couleur()
    with Sheets("MAPPING").Range("A:Z").formatconditions.add(xlCellValue, xlEqual, "X")
        .interior.colorindex()
        


End Sub

Sub recup()


    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim FolderPath As String
    Dim FileName As String
    Dim twb As Workbook
    


        Set twb = ActiveWorkbook
        
        FolderPath = Sheets(1).Range("F2").Value
        FileName = Dir(FolderPath & "\*.xls*")
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(FolderPath & "\" & FileName)



        wb.Sheets(1).Copy after:=twb.Sheets(1)
        twb.Sheets(2).Name = "Export"



End Sub

Sub trier()

    Dim der As Integer
    Dim n As Integer
    Dim ann As String
    Dim div As String
    Dim wsdiv As Worksheet

    der = Sheets(2).Range("B" & Rows.Count).End(xlUp).Row
    For n = 1 To der
        ann = Sheets(2).Range("Q" & n).Value
        If ann = Sheets(1).Range("F4").Value Then
        
            div = Sheets(2).Range("D" & n).Value
            
            On Error Resume Next
            Set wsdiv = ThisWorkbook.Sheets(div)
            On Error GoTo 0

        
            If wsdiv Is Nothing Then
                Set wsdiv = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                wsdiv.Name = div
                Sheets(2).Range("2:2").Copy Sheets(div).Range("2:2")
                
                
                
            End If
            Set wsdiv = Nothing
            fin = Sheets(div).Range("B" & Rows.Count).End(xlUp).Row + 1
            
            Sheets(2).Range("A" & n).EntireRow.Copy Sheets(div).Range("A" & fin)
            
            
            
        End If
        
    
    Next n

End Sub

Sub clean()

    For i = 1 To Sheets.Count
        If Sheets(i).Name <> "Macro" Then
            If Sheets(i).Name <> "STATS" Then
                If Sheets(i).Name <> "calculs" Then
                    Sheets(i).Delete
                End If
            End If
        End If
    Next i
    
        



End Sub

Sub janvier()
    
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I20").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G20").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I19").Value)
        Sheets("Calculs").Range("J20").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G20").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J19").Value)
        Sheets("Calculs").Range("K20").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G20").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K19").Value)
        Sheets("Calculs").Range("L20").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G20").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L19").Value)
        Sheets("Calculs").Range("M20").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G20").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M19").Value)
        Sheets("Calculs").Range("N20").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G20").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N19").Value)
        Sheets("Calculs").Range("O20").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G20").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O19").Value)
        
        

End Sub

Sub fevrier()

    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        
        Sheets("Calculs").Range("I28").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G28").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I27").Value)
        Sheets("Calculs").Range("J28").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G28").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J27").Value)
        Sheets("Calculs").Range("K28").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G28").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K27").Value)
        Sheets("Calculs").Range("L28").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G28").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L27").Value)
        Sheets("Calculs").Range("M28").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G28").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M27").Value)
        Sheets("Calculs").Range("N28").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G28").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N27").Value)
        Sheets("Calculs").Range("O28").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G28").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O27").Value)
        
   
End Sub

Sub mars()
    
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I36").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G36").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I35").Value)
        Sheets("Calculs").Range("J36").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G36").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J35").Value)
        Sheets("Calculs").Range("K36").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G36").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K35").Value)
        Sheets("Calculs").Range("L36").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G36").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L35").Value)
        Sheets("Calculs").Range("M36").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G36").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M35").Value)
        Sheets("Calculs").Range("N36").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G36").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N35").Value)
        Sheets("Calculs").Range("O36").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G36").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O35").Value)
End Sub

Sub avril()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I44").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G44").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I43").Value)
        Sheets("Calculs").Range("J44").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G44").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J43").Value)
        Sheets("Calculs").Range("K44").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G44").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K43").Value)
        Sheets("Calculs").Range("L44").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G44").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L43").Value)
        Sheets("Calculs").Range("M44").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G44").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M43").Value)
        Sheets("Calculs").Range("N44").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G44").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N43").Value)
        Sheets("Calculs").Range("O44").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G44").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O43").Value)

End Sub

Sub mai()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I52").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G52").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I51").Value)
        Sheets("Calculs").Range("J52").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G52").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J51").Value)
        Sheets("Calculs").Range("K52").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G52").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K51").Value)
        Sheets("Calculs").Range("L52").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G52").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L51").Value)
        Sheets("Calculs").Range("M52").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G52").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M51").Value)
        Sheets("Calculs").Range("N52").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G52").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N51").Value)
        Sheets("Calculs").Range("O52").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G52").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O51").Value)

End Sub

Sub juin()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I60").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G60").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I59").Value)
        Sheets("Calculs").Range("J60").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G60").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J59").Value)
        Sheets("Calculs").Range("K60").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G60").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K59").Value)
        Sheets("Calculs").Range("L60").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G60").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L59").Value)
        Sheets("Calculs").Range("M60").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G60").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M59").Value)
        Sheets("Calculs").Range("N60").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G60").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N59").Value)
        Sheets("Calculs").Range("O60").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G60").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O59").Value)

End Sub

Sub juillet()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I68").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G68").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I67").Value)
        Sheets("Calculs").Range("J68").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G68").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J67").Value)
        Sheets("Calculs").Range("K68").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G68").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K67").Value)
        Sheets("Calculs").Range("L68").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G68").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L67").Value)
        Sheets("Calculs").Range("M68").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G68").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M67").Value)
        Sheets("Calculs").Range("N68").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G68").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N67").Value)
        Sheets("Calculs").Range("O68").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G68").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O67").Value)

End Sub

Sub aout()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I76").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G76").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I75").Value)
        Sheets("Calculs").Range("J76").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G76").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J75").Value)
        Sheets("Calculs").Range("K76").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G76").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K75").Value)
        Sheets("Calculs").Range("L76").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G76").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L75").Value)
        Sheets("Calculs").Range("M76").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G76").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M75").Value)
        Sheets("Calculs").Range("N76").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G76").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N75").Value)
        Sheets("Calculs").Range("O76").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G76").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O75").Value)

End Sub

Sub septembre()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I84").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G84").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I83").Value)
        Sheets("Calculs").Range("J84").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G84").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J83").Value)
        Sheets("Calculs").Range("K84").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G84").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K83").Value)
        Sheets("Calculs").Range("L84").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G84").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L83").Value)
        Sheets("Calculs").Range("M84").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G84").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M83").Value)
        Sheets("Calculs").Range("N84").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G84").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N83").Value)
        Sheets("Calculs").Range("O84").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G84").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O83").Value)

End Sub

Sub octobre()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I92").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G92").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I91").Value)
        Sheets("Calculs").Range("J92").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G92").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J91").Value)
        Sheets("Calculs").Range("K92").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G92").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K91").Value)
        Sheets("Calculs").Range("L92").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G92").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L91").Value)
        Sheets("Calculs").Range("M92").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G92").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M91").Value)
        Sheets("Calculs").Range("N92").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G92").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N91").Value)
        Sheets("Calculs").Range("O92").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G92").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O91").Value)

End Sub

Sub novembre()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I100").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G100").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I100").Value)
        Sheets("Calculs").Range("J100").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G100").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J100").Value)
        Sheets("Calculs").Range("K100").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G100").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K100").Value)
        Sheets("Calculs").Range("L100").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G100").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L100").Value)
        Sheets("Calculs").Range("M100").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G100").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M100").Value)
        Sheets("Calculs").Range("N100").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G100").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N100").Value)
        Sheets("Calculs").Range("O100").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G100").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O100").Value)

End Sub

Sub decembre()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)
        
        

        Sheets("Calculs").Range("I108").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G108").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("I108").Value)
        Sheets("Calculs").Range("J108").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G108").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("J108").Value)
        Sheets("Calculs").Range("K108").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G108").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("K108").Value)
        Sheets("Calculs").Range("L108").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G108").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("L108").Value)
        Sheets("Calculs").Range("M108").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G108").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("M108").Value)
        Sheets("Calculs").Range("N108").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G108").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("N108").Value)
        Sheets("Calculs").Range("O108").Value = WorksheetFunction.CountIfs(Sheets(div).Range("P:P"), Sheets("Calculs").Range("G108").Value, Sheets(div).Range("S:S"), Sheets("Calculs").Range("O108").Value)

End Sub

Sub total()
        
    Dim div As String
    Dim i As Integer
    Dim n As Integer
    Dim der As Integer
    
    
        div = Sheets("STATS").DropDowns("Zone combinée 1").List(Sheets("STATS").DropDowns("Zone combinée 1").ListIndex)

        Sheets("Calculs").Range("I116").Value = WorksheetFunction.CountIfs(Sheets(div).Range("S:S"), Sheets("Calculs").Range("I115").Value)
        Sheets("Calculs").Range("J116").Value = WorksheetFunction.CountIfs(Sheets(div).Range("S:S"), Sheets("Calculs").Range("J115").Value)
        Sheets("Calculs").Range("K116").Value = WorksheetFunction.CountIfs(Sheets(div).Range("S:S"), Sheets("Calculs").Range("K115").Value)
        Sheets("Calculs").Range("L116").Value = WorksheetFunction.CountIfs(Sheets(div).Range("S:S"), Sheets("Calculs").Range("L115").Value)
        Sheets("Calculs").Range("M116").Value = WorksheetFunction.CountIfs(Sheets(div).Range("S:S"), Sheets("Calculs").Range("M115").Value)
        Sheets("Calculs").Range("N116").Value = WorksheetFunction.CountIfs(Sheets(div).Range("S:S"), Sheets("Calculs").Range("N115").Value)
        Sheets("Calculs").Range("O116").Value = WorksheetFunction.CountIfs(Sheets(div).Range("S:S"), Sheets("Calculs").Range("O115").Value)

End Sub

Sub magic()


    Dim n As Integer
    Dim a As Integer
    Dim der As Integer
    Dim job As String
    Dim classif As String
    Dim b As Integer
    Dim dercls As Integer
    Dim cazclassf As Range
    Dim lgn As Integer
    Dim coln As Integer
    Dim cazjob As Range
    Dim lgnjob As Integer
    Dim coljob As Integer


    dercls = Sheets(2).Range("A" & Rows.Count).End(xlUp).Row
    der = Sheets("DATA BASE").Range("A" & Rows.Count).End(xlUp).Row
    div = Sheets("MAPPING").DropDowns("Zone combinée 4").List(Sheets("MAPPING").DropDowns("Zone combinée 4").ListIndex)
    fil = Sheets("MAPPING").DropDowns("Zone combinée 5").List(Sheets("MAPPING").DropDowns("Zone combinée 5").ListIndex)

    For n = 2 To dercls
        i = n
        classif = Sheets(2).Range("A" & n)
        Sheets("MAPPING").Range("A" & n).Value = classif

    Next n

    b = 2



    If div = "Toutes Divisions" Then
        If fil = "Toutes Filieres" Then
            b = 2
            For a = 2 To der


                job = Sheets("DATA BASE").Range("A" & a).Value
                classif = Sheets("DATA BASE").Range("B" & a)


                Set cazjob = Sheets("MAPPING").Range(Cells(dercls + 1, 2), Cells(dercls + 1, b)).Find(job)
                If cazjob Is Nothing Then
                    Sheets("MAPPING").Range(Cells(dercls + 1, b), Cells(dercls + 1, b)).Value = job
                    lgnjob = dercls + 1
                    coljob = b
                    b = b + 1
                Else
                    lgnjob = cazjob.Row
                    coljob = cazjob.Column


                End If
                Set cazclassf = Sheets("MAPPING").Range(Cells(2, 1), Cells(dercls, 2)).Find(classif)
                lgn = cazclassf.Row
                coln = cazclassf.Column
                Sheets("MAPPING").Range(Cells(lgn, coljob), Cells(lgn, coljob)).Value = "X"



            Next a
        End If
        If fil <> "Toutes Filieres" Then
             For a = 2 To der
              If Sheets("DATA BASE").Range("D" & a).Value = fil Then

                job = Sheets("DATA BASE").Range("A" & a).Value
                classif = Sheets("DATA BASE").Range("B" & a)


                Set cazjob = Sheets("MAPPING").Range(Cells(dercls + 1, 2), Cells(dercls + 1, b)).Find(job)
                If cazjob Is Nothing Then
                    Sheets("MAPPING").Range(Cells(dercls + 1, b), Cells(dercls + 1, b)).Value = job
                    lgnjob = dercls + 1
                    coljob = b
                    b = b + 1
                Else
                    lgnjob = cazjob.Row
                    coljob = cazjob.Column


                End If
                Set cazclassf = Sheets("MAPPING").Range(Cells(2, 1), Cells(dercls, 2)).Find(classif)
                lgn = cazclassf.Row
                coln = cazclassf.Column
                Sheets("MAPPING").Range(Cells(lgn, coljob), Cells(lgn, coljob)).Value = "X"
               End If
            Next a
        End If
    End If
    If div <> "Toutes Divisions" Then
        If fil = "Toutes Filieres" Then
            b = 2
            For a = 2 To der
              If Sheets("DATA BASE").Range("C" & a).Value = div Then

                job = Sheets("DATA BASE").Range("A" & a).Value
                classif = Sheets("DATA BASE").Range("B" & a)


                Set cazjob = Sheets("MAPPING").Range(Cells(dercls + 1, 2), Cells(dercls + 1, b)).Find(job)
                If cazjob Is Nothing Then
                    Sheets("MAPPING").Range(Cells(dercls + 1, b), Cells(dercls + 1, b)).Value = job
                    lgnjob = dercls + 1
                    coljob = b
                    b = b + 1
                Else
                    lgnjob = cazjob.Row
                    coljob = cazjob.Column


                End If
                Set cazclassf = Sheets("MAPPING").Range(Cells(2, 1), Cells(dercls, 2)).Find(classif)
                lgn = cazclassf.Row
                coln = cazclassf.Column
                Sheets("MAPPING").Range(Cells(lgn, coljob), Cells(lgn, coljob)).Value = "X"
               End If
            Next a
        End If
        If fil <> "Toutes Filieres" Then
             For a = 2 To der
              If Sheets("DATA BASE").Range("D" & a).Value = fil and Sheets("DATA BASE").Range("C" & a).Value = div Then

                job = Sheets("DATA BASE").Range("A" & a).Value
                classif = Sheets("DATA BASE").Range("B" & a)


                Set cazjob = Sheets("MAPPING").Range(Cells(dercls + 1, 2), Cells(dercls + 1, b)).Find(job)
                If cazjob Is Nothing Then
                    Sheets("MAPPING").Range(Cells(dercls + 1, b), Cells(dercls + 1, b)).Value = job
                    lgnjob = dercls + 1
                    coljob = b
                    b = b + 1
                Else
                    lgnjob = cazjob.Row
                    coljob = cazjob.Column


                End If
                Set cazclassf = Sheets("MAPPING").Range(Cells(2, 1), Cells(dercls, 2)).Find(classif)
                lgn = cazclassf.Row
                coln = cazclassf.Column
                Sheets("MAPPING").Range(Cells(lgn, coljob), Cells(lgn, coljob)).Value = "X"
               End If
            Next a
        End If
    End If

End Sub

sub ca()
    Dim oldwb As Workbook
    Dim newwb As Workbook
    Dim der As Integer
    Dim i As Integer
    Dim a As Integer
    Dim mec As String
    dim nom as String
    dim prenon as String
    Dim div As String
    Dim FolderPath As String
    Dim FileName As String
    Dim twb As Workbook
    Dim out As String
    
    Set newwb = Workbooks.Open(Sheets(1).Range("F3").Value & "\" & "neww")
    Set twb = ActiveWorkbook
        out = Sheets(1).Range("F3").Value
        FolderPath = Sheets(1).Range("F2").Value
        FileName = Dir(FolderPath & "\*.xls*")
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(FolderPath & "\" & FileName)     

     Do While FileName <> ""
        nom=wb.Sheets("TRAME VIERGE").Range("C8").Value 
        prenom=wb.Sheets("TRAME VIERGE").Range("M8").Value
        mec = nom & " " & prenom
        div==wb.DropDowns("Zone combinée 140").List(wb.DropDowns("Zone combinée 140").ListIndex)
        newwb.sheets(1).range("C9").value=div
        newwb.Sheets(1).Range("C8").Value = wb.Sheets("TRAME VIERGE").Range("C8").Value
        newwb.Sheets(1).Range("Q6").Value = wb.Sheets("TRAME VIERGE").Range("I6").Value
        newwb.Sheets(1).Range("M8").Value = wb.Sheets("TRAME VIERGE").Range("M8").Value
        newwb.Sheets(1).Range("H9").Value = wb.Sheets("TRAME VIERGE").Range("H9").Value
        newwb.Sheets(1).Range("R9").Value = wb.Sheets("TRAME VIERGE").Range("R9").Value
        newwb.Sheets(1).Range("C11").Value = wb.Sheets("TRAME VIERGE").Range("C11").Value
        newwb.Sheets(1).Range("I11").Value = wb.Sheets("TRAME VIERGE").Range("I11").Value
        newwb.Sheets(1).Range("P11").Value = wb.Sheets("TRAME VIERGE").Range("P11").Value
        newwb.Sheets(1).Range("A16").Value = wb.Sheets("TRAME VIERGE").Range("A16").Value
        newwb.Sheets(1).Range("A17").Value = wb.Sheets("TRAME VIERGE").Range("A17").Value
        newwb.Sheets(1).Range("A18").Value = wb.Sheets("TRAME VIERGE").Range("A18").Value
        newwb.Sheets(1).Range("A19").Value = wb.Sheets("TRAME VIERGE").Range("A19").Value
        newwb.Sheets(1).Range("A22").Value = wb.Sheets("TRAME VIERGE").Range("A22").Value
        newwb.Sheets(1).Range("A23").Value = wb.Sheets("TRAME VIERGE").Range("A23").Value
        newwb.Sheets(1).Range("A24").Value = wb.Sheets("TRAME VIERGE").Range("A17").Value
        newwb.Sheets(1).Range("A25").Value = wb.Sheets("TRAME VIERGE").Range("A25").Value
        newwb.Sheets(1).Range("A26").Value = wb.Sheets("TRAME VIERGE").Range("A26").Value
        newwb.Sheets(1).Range("A27").Value = wb.Sheets("TRAME VIERGE").Range("A27").Value
        newwb.Sheets(1).Range("A28").Value = wb.Sheets("TRAME VIERGE").Range("A28").Value
        newwb.Sheets(1).Range("A29").Value = wb.Sheets("TRAME VIERGE").Range("A29").Value
        newwb.Sheets(1).Range("A30").Value = wb.Sheets("TRAME VIERGE").Range("A30").Value
        newwb.Sheets(1).Range("A31").Value = wb.Sheets("TRAME VIERGE").Range("A31").Value
        newwb.Sheets(1).Range("A32").Value = wb.Sheets("TRAME VIERGE").Range("A32").Value
        newwb.Sheets(1).Range("A33").Value = wb.Sheets("TRAME VIERGE").Range("A33").Value
        newwb.Sheets(1).Range("A34").Value = wb.Sheets("TRAME VIERGE").Range("A34").Value
        newwb.Sheets(1).Range("A37").Value = wb.Sheets("TRAME VIERGE").Range("B45").Value
        newwb.Sheets(1).Range("A38").Value = wb.Sheets("TRAME VIERGE").Range("B46").Value
        newwb.Sheets(1).Range("A39").Value = wb.Sheets("TRAME VIERGE").Range("B47").Value
        newwb.Sheets(1).Range("A40").Value = wb.Sheets("TRAME VIERGE").Range("B48").Value


        newwb.Sheets(1).Range("F70").Value = wb.Sheets("TRAME VIERGE").Range("F70").Value
        newwb.Sheets(1).Range("F71").Value = wb.Sheets("TRAME VIERGE").Range("F71").Value
        newwb.Sheets(1).Range("F72").Value = wb.Sheets("TRAME VIERGE").Range("F72").Value
        newwb.Sheets(1).Range("F73").Value = wb.Sheets("TRAME VIERGE").Range("F73").Value
        newwb.Sheets(1).Range("F74").Value = wb.Sheets("TRAME VIERGE").Range("F74").Value
        newwb.Sheets(1).Range("A85").Value = wb.Sheets("TRAME VIERGE").Range("A85").Value
        newwb.Sheets(1).Range("A86").Value = wb.Sheets("TRAME VIERGE").Range("A86").Value
        newwb.Sheets(1).Range("A87").Value = wb.Sheets("TRAME VIERGE").Range("A87").Value
        newwb.Sheets(1).Range("A88").Value = wb.Sheets("TRAME VIERGE").Range("A88").Value
        newwb.Sheets(1).Range("A89").Value = wb.Sheets("TRAME VIERGE").Range("A89").Value
        newwb.Sheets(1).Range("F85").Value = wb.Sheets("TRAME VIERGE").Range("F85").Value
        newwb.Sheets(1).Range("F86").Value = wb.Sheets("TRAME VIERGE").Range("F86").Value
        newwb.Sheets(1).Range("F87").Value = wb.Sheets("TRAME VIERGE").Range("F87").Value
        newwb.Sheets(1).Range("F88").Value = wb.Sheets("TRAME VIERGE").Range("F88").Value
        newwb.Sheets(1).Range("F89").Value = wb.Sheets("TRAME VIERGE").Range("F89").Value
        newwb.Sheets(1).Range("N85").Value = wb.Sheets("TRAME VIERGE").Range("N85").Value
        newwb.Sheets(1).Range("N86").Value = wb.Sheets("TRAME VIERGE").Range("N86").Value
        newwb.Sheets(1).Range("N87").Value = wb.Sheets("TRAME VIERGE").Range("N87").Value
        newwb.Sheets(1).Range("N88").Value = wb.Sheets("TRAME VIERGE").Range("N88").Value
        newwb.Sheets(1).Range("N89").Value = wb.Sheets("TRAME VIERGE").Range("N89").Value
        newwb.Sheets(1).Range("Q85").Value = wb.Sheets("TRAME VIERGE").Range("Q85").Value
        newwb.Sheets(1).Range("Q86").Value = wb.Sheets("TRAME VIERGE").Range("Q86").Value
        newwb.Sheets(1).Range("Q87").Value = wb.Sheets("TRAME VIERGE").Range("Q87").Value
        newwb.Sheets(1).Range("Q88").Value = wb.Sheets("TRAME VIERGE").Range("Q88").Value
        newwb.Sheets(1).Range("Q89").Value = wb.Sheets("TRAME VIERGE").Range("Q89").Value


        newwb.Sheets(1).Range("A92").Value = wb.Sheets("TRAME VIERGE").Range("A92").Value
        newwb.Sheets(1).Range("A93").Value = wb.Sheets("TRAME VIERGE").Range("A93").Value
        newwb.Sheets(1).Range("A94").Value = wb.Sheets("TRAME VIERGE").Range("A94").Value
        newwb.Sheets(1).Range("A95").Value = wb.Sheets("TRAME VIERGE").Range("A95").Value
        newwb.Sheets(1).Range("A96").Value = wb.Sheets("TRAME VIERGE").Range("A96").Value
        newwb.Sheets(1).Range("F92").Value = wb.Sheets("TRAME VIERGE").Range("F92").Value
        newwb.Sheets(1).Range("F93").Value = wb.Sheets("TRAME VIERGE").Range("F93").Value
        newwb.Sheets(1).Range("F94").Value = wb.Sheets("TRAME VIERGE").Range("F94").Value
        newwb.Sheets(1).Range("F95").Value = wb.Sheets("TRAME VIERGE").Range("F95").Value
        newwb.Sheets(1).Range("F96").Value = wb.Sheets("TRAME VIERGE").Range("F96").Value
        newwb.Sheets(1).Range("N92").Value = wb.Sheets("TRAME VIERGE").Range("N92").Value
        newwb.Sheets(1).Range("N93").Value = wb.Sheets("TRAME VIERGE").Range("N93").Value
        newwb.Sheets(1).Range("N94").Value = wb.Sheets("TRAME VIERGE").Range("N94").Value
        newwb.Sheets(1).Range("N95").Value = wb.Sheets("TRAME VIERGE").Range("N95").Value
        newwb.Sheets(1).Range("N96").Value = wb.Sheets("TRAME VIERGE").Range("N96").Value
        newwb.Sheets(1).Range("Q92").Value = wb.Sheets("TRAME VIERGE").Range("Q92").Value
        newwb.Sheets(1).Range("Q93").Value = wb.Sheets("TRAME VIERGE").Range("Q93").Value
        newwb.Sheets(1).Range("Q94").Value = wb.Sheets("TRAME VIERGE").Range("Q94").Value
        newwb.Sheets(1).Range("Q95").Value = wb.Sheets("TRAME VIERGE").Range("Q95").Value
        newwb.Sheets(1).Range("Q96").Value = wb.Sheets("TRAME VIERGE").Range("Q96").Value


            newwb.saveas(out & "\" & div & mec & ".xlsx")

    Name FolderPath & "\" & FileName As out & "\" & div & " " & uniqueID & ".xlsm"
        
        FileName = Dir
     Loop





End Sub

Sub Bouger_ok(Filename)
    

    Dim FichierOriginal As String
    Dim FichierDeplace As String
    
    FichierOriginal = "C:\Users\Alexis.Soucaze\OneDrive - BASTIDE MEDICAL\Bureau\Test xl\lnl\ent an et pro\vieux" & Filename
    FichierDeplace = "C:\Users\Alexis.Soucaze\OneDrive - BASTIDE MEDICAL\Bureau\Test xl\lnl\ent an et pro\ok" & Filename
    
    Name FichierOriginal As FichierDeplace
    
End Sub

Sub Bouger_echec(Filename, div)

    

    Dim FichierOriginal As String
    Dim FichierDeplace As String
    
    FichierOriginal = "C:\Users\Alexis.Soucaze\OneDrive - BASTIDE MEDICAL\Bureau\Test xl\lnl\ent an et pro\vieux" & Filename
    FichierDeplace = "\\srvfichier\drh\WINDOWS\drh\ENTRETIENS ANNUELS et PROFESSIONNELS\2023-2024\CONSOLIDATION\Entretien en erreur\" & "Erreur " & Filename
    

    Name FichierOriginal As FichierDeplace
End Sub

sub tt()
    If Not div = "Collectivités - HAD" or "MAD
    " "Plateforme
    " "Respiratoire"


    If Not div = "Collectivités - HAD" or  Not div = "MAD" or  Not div = "Plateforme" or Not div = "Respiratoire" or Not div = "Diabète" or Not div = "Nutrition -Perfusion" or Not div = "Siège" or Not div = "SUC" or Not div = "Anissa Pâtisserie" Then



            newwb.Sheets(1).Range("C9").Value = div
            newwb.Sheets(1).Range("C8").Value = wb.Sheets(1).Range("C8").Value
            newwb.Sheets(1).Range("Q6").Value = wb.Sheets(1).Range("I6").Value
            newwb.Sheets(1).Range("M8").Value = wb.Sheets(1).Range("M8").Value
            newwb.Sheets(1).Range("H9").Value = wb.Sheets(1).Range("H9").Value
            newwb.Sheets(1).Range("R9").Value = wb.Sheets(1).Range("R9").Value
            newwb.Sheets(1).Range("C11").Value = wb.Sheets(1).Range("C11").Value
            newwb.Sheets(1).Range("I11").Value = wb.Sheets(1).Range("I11").Value
            newwb.Sheets(1).Range("P11").Value = wb.Sheets(1).Range("P11").Value
            newwb.Sheets(1).Range("A16").Value = wb.Sheets(1).Range("A16").Value
            newwb.Sheets(1).Range("A17").Value = wb.Sheets(1).Range("A17").Value
            newwb.Sheets(1).Range("A18").Value = wb.Sheets(1).Range("A18").Value
            newwb.Sheets(1).Range("A19").Value = wb.Sheets(1).Range("A19").Value
            newwb.Sheets(1).Range("A26").Value = wb.Sheets(1).Range("A22").Value
            newwb.Sheets(1).Range("A27").Value = wb.Sheets(1).Range("A23").Value
            newwb.Sheets(1).Range("A28").Value = wb.Sheets(1).Range("A17").Value
            newwb.Sheets(1).Range("A29").Value = wb.Sheets(1).Range("A25").Value
            newwb.Sheets(1).Range("A30").Value = wb.Sheets(1).Range("A26").Value
            newwb.Sheets(1).Range("A31").Value = wb.Sheets(1).Range("A27").Value
            newwb.Sheets(1).Range("A32").Value = wb.Sheets(1).Range("A28").Value
            newwb.Sheets(1).Range("A33").Value = wb.Sheets(1).Range("A29").Value
            newwb.Sheets(1).Range("A34").Value = wb.Sheets(1).Range("A30").Value
            newwb.Sheets(1).Range("A35").Value = wb.Sheets(1).Range("A31").Value
            newwb.Sheets(1).Range("A36").Value = wb.Sheets(1).Range("A32").Value
            newwb.Sheets(1).Range("A37").Value = wb.Sheets(1).Range("A33").Value
            newwb.Sheets(1).Range("A38").Value = wb.Sheets(1).Range("A34").Value

            newwb.Sheets(1).Range("A41").Value = wb.Sheets(1).Range("B45").Value
            newwb.Sheets(1).Range("A42").Value = wb.Sheets(1).Range("B46").Value
            newwb.Sheets(1).Range("A43").Value = wb.Sheets(1).Range("B47").Value
            newwb.Sheets(1).Range("A44").Value = wb.Sheets(1).Range("B48").Value


            newwb.Sheets(1).Range("F74").Value = wb.Sheets(1).Range("F70").Value
            newwb.Sheets(1).Range("F75").Value = wb.Sheets(1).Range("F71").Value
            newwb.Sheets(1).Range("F76").Value = wb.Sheets(1).Range("F72").Value
            newwb.Sheets(1).Range("F77").Value = wb.Sheets(1).Range("F73").Value
            newwb.Sheets(1).Range("F78").Value = wb.Sheets(1).Range("F74").Value


            newwb.Sheets(1).Range("F80").Value = wb.Sheets(1).Range("F76").Value
            newwb.Sheets(1).Range("F82").Value = wb.Sheets(1).Range("F78").Value
            newwb.Sheets(1).Range("F84").Value = wb.Sheets(1).Range("F80").Value


            newwb.Sheets(1).Range("A89").Value = wb.Sheets(1).Range("A85").Value
            newwb.Sheets(1).Range("A90").Value = wb.Sheets(1).Range("A86").Value
            newwb.Sheets(1).Range("A91").Value = wb.Sheets(1).Range("A87").Value
            newwb.Sheets(1).Range("A92").Value = wb.Sheets(1).Range("A88").Value
            newwb.Sheets(1).Range("A93").Value = wb.Sheets(1).Range("A89").Value
            newwb.Sheets(1).Range("F89").Value = wb.Sheets(1).Range("F85").Value
            newwb.Sheets(1).Range("F90").Value = wb.Sheets(1).Range("F86").Value
            newwb.Sheets(1).Range("F91").Value = wb.Sheets(1).Range("F87").Value
            newwb.Sheets(1).Range("F92").Value = wb.Sheets(1).Range("F88").Value
            newwb.Sheets(1).Range("F93").Value = wb.Sheets(1).Range("F89").Value
            newwb.Sheets(1).Range("N89").Value = wb.Sheets(1).Range("N85").Value
            newwb.Sheets(1).Range("N90").Value = wb.Sheets(1).Range("N86").Value
            newwb.Sheets(1).Range("N91").Value = wb.Sheets(1).Range("N87").Value
            newwb.Sheets(1).Range("N92").Value = wb.Sheets(1).Range("N88").Value
            newwb.Sheets(1).Range("N93").Value = wb.Sheets(1).Range("N89").Value
            newwb.Sheets(1).Range("Q89").Value = wb.Sheets(1).Range("Q85").Value
            newwb.Sheets(1).Range("Q90").Value = wb.Sheets(1).Range("Q86").Value
            newwb.Sheets(1).Range("Q91").Value = wb.Sheets(1).Range("Q87").Value
            newwb.Sheets(1).Range("Q92").Value = wb.Sheets(1).Range("Q88").Value
            newwb.Sheets(1).Range("Q93").Value = wb.Sheets(1).Range("Q89").Value


            newwb.Sheets(1).Range("A96").Value = wb.Sheets(1).Range("A92").Value
            newwb.Sheets(1).Range("A97").Value = wb.Sheets(1).Range("A93").Value
            newwb.Sheets(1).Range("A98").Value = wb.Sheets(1).Range("A94").Value
            newwb.Sheets(1).Range("A99").Value = wb.Sheets(1).Range("A95").Value
            newwb.Sheets(1).Range("A100").Value = wb.Sheets(1).Range("A96").Value
            newwb.Sheets(1).Range("F96").Value = wb.Sheets(1).Range("F92").Value
            newwb.Sheets(1).Range("F97").Value = wb.Sheets(1).Range("F93").Value
            newwb.Sheets(1).Range("F98").Value = wb.Sheets(1).Range("F94").Value
            newwb.Sheets(1).Range("F99").Value = wb.Sheets(1).Range("F95").Value
            newwb.Sheets(1).Range("F100").Value = wb.Sheets(1).Range("F96").Value
            newwb.Sheets(1).Range("N96").Value = wb.Sheets(1).Range("N92").Value
            newwb.Sheets(1).Range("N97").Value = wb.Sheets(1).Range("N93").Value
            newwb.Sheets(1).Range("N98").Value = wb.Sheets(1).Range("N94").Value
            newwb.Sheets(1).Range("N99").Value = wb.Sheets(1).Range("N95").Value
            newwb.Sheets(1).Range("N100").Value = wb.Sheets(1).Range("N96").Value
            newwb.Sheets(1).Range("Q96").Value = wb.Sheets(1).Range("Q92").Value
            newwb.Sheets(1).Range("Q97").Value = wb.Sheets(1).Range("Q93").Value
            newwb.Sheets(1).Range("Q98").Value = wb.Sheets(1).Range("Q94").Value
            newwb.Sheets(1).Range("Q99").Value = wb.Sheets(1).Range("Q95").Value
            newwb.Sheets(1).Range("Q100").Value = wb.Sheets(1).Range("Q96").Value

End Sub


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

