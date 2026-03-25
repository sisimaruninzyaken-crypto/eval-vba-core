Attribute VB_Name = "modMMTLayout"
Public Sub Resize_MMTChildHost_ToPage()
    
    Dim pg As Object, host As Object, child As Object

    Set pg = GetMMTPage(frmEval)
    If pg Is Nothing Then Exit Sub

    Set host = GetMMTHost(pg)
    Set child = GetMMTChildTabs(pg, host)
    If host Is Nothing Or child Is Nothing Then Exit Sub

    host.Width = pg.InsideWidth - 12
    host.Height = pg.InsideHeight - 12

    child.Left = 0
    child.top = 0
    child.Width = host.InsideWidth
    child.Height = host.InsideHeight
End Sub

