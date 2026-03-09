Attribute VB_Name = "modMMTLayout"
Public Sub Resize_MMTChildHost_ToPage()
    
    Debug.Print "[MMTResize] pgInside=", frmEval.controls("mpPhys").Pages(1).InsideWidth, frmEval.controls("mpPhys").Pages(1).InsideHeight

    
    Dim mp As Object, pg As Object, host As Object, child As Object

    Set mp = frmEval.controls("mpPhys")
    mp.value = 1
    DoEvents

    Set pg = mp.Pages(1)
    Set host = GetMMTHost(pg)
    Set child = GetMMTChildTabs(pg, host)
    If host Is Nothing Or child Is Nothing Then Exit Sub

    host.Width = pg.InsideWidth - 12
    host.Height = pg.InsideHeight - 12

    child.Left = 0
    child.Top = 0
    child.Width = host.InsideWidth
    child.Height = host.InsideHeight
End Sub

