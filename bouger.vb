Sub Bouger_ok()
 
Dim FolderPath As String
Dim Filename As String
Dim FichierOriginal As String
Dim FichierDeplace As String

    FolderPath = ThisWorkbook.Sheets(1).Range("C3").Value
    Filename = Dir(FolderPath & "\*.xls*") 

FichierOriginal = "\\srvfichier\drh\WINDOWS\drh\ENTRETIENS ANNUELS et PROFESSIONNELS\2024-2025\ENTRETIENS A DEPOSER ICI\Entretiens annuels\" & Filename
FichierDeplace = "\\srvfichier\drh\WINDOWS\drh\ENTRETIENS ANNUELS et PROFESSIONNELS\2024-2025\CONSOLIDATION\EA\enregistrement_ok\" & Filename& ".xlsx"
 
Name FichierOriginal As FichierDeplace
 
End Sub

