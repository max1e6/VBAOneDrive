Attribute VB_Name = "basOneDrive"
Public Sub TestOneDrive()

    Dim poOD As New COneDrive
    Dim iRow As Integer
    Dim v As Variant
    
    ThisWorkbook.Worksheets(1).Cells.Clear
    
    poOD.URI = ThisWorkbook.Path
    
    iRow = 18
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "URI"
    ThisWorkbook.Worksheets(1).Cells(iRow, 3) = poOD.URI
    
    iRow = iRow + 1
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "Is URI"
    ThisWorkbook.Worksheets(1).Cells(iRow, 3) = poOD.IsURI
    
    iRow = iRow + 1
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "OneDrive Type"
    ThisWorkbook.Worksheets(1).Cells(iRow, 3) = poOD.OneDriveType
     
    iRow = iRow + 1
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "Local Path"
    ThisWorkbook.Worksheets(1).Cells(iRow, 3) = poOD.LocalPath
    
    iRow = iRow + 2
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "CID"
    ThisWorkbook.Worksheets(1).Cells(iRow, 3) = poOD.OneDriveCID
    
    iRow = iRow + 1
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "Consumer Path"
    ThisWorkbook.Worksheets(1).Cells(iRow, 3) = poOD.OneDriveConsumerPath
    
    iRow = iRow + 1
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "Commercial Path"
    ThisWorkbook.Worksheets(1).Cells(iRow, 3) = poOD.OneDriveCommercialPath
    
    iRow = iRow + 1
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "OneDrive URI"
    ThisWorkbook.Worksheets(1).Cells(iRow, 3) = poOD.OneDriveURI
    
    iRow = iRow + 1
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "Teams URI"
    ThisWorkbook.Worksheets(1).Cells(iRow, 3) = poOD.TeamsURI

    iRow = iRow + 2
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "Tenants"
    For Each v In poOD.Tenants
      ThisWorkbook.Worksheets(1).Cells(iRow, 3) = v
      iRow = iRow + 1
    Next
    
    iRow = iRow + 1
    ThisWorkbook.Worksheets(1).Cells(iRow, 2) = "Channels"
    For Each v In poOD.Channels
      ThisWorkbook.Worksheets(1).Cells(iRow, 3) = v
      iRow = iRow + 1
    Next
   
    ThisWorkbook.Worksheets(1).Columns(2).Font.Bold = True
    
End Sub
