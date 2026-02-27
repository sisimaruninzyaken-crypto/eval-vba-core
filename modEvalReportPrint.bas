Attribute VB_Name = "modEvalReportPrint"
Option Explicit


Private Sub WalkContainer(ByVal cont As Object, ByRef maxBottom As Double)
    
   '--- MultiPage は Controls を持たない（Pages を掘る）
If TypeName(cont) = "MultiPage" Then
    Dim p As MSForms.Page
    For Each p In cont.Pages
        WalkContainer p, maxBottom
    Next p
    Exit Sub
End If
 
    
    
    
    Dim c As MSForms.Control

    Static lastMaxInfo As String


    For Each c In cont.Controls

    Dim isContainer As Boolean
    isContainer = (TypeOf c Is MSForms.Frame) _
                  Or (TypeOf c Is MSForms.MultiPage) _
                  Or (TypeOf c Is MSForms.Page)

    ' 葉（入力部品など）だけで maxBottom を更新する
    If c.Visible Then
        If Not isContainer Then
            If c.Top + c.Height > maxBottom Then maxBottom = c.Top + c.Height: lastMaxInfo = TypeName(c) & "  " & c.name & "  Bottom=" & (c.Top + c.Height)
            If c.Top + c.Height > maxBottom Then lastMaxInfo = TypeName(c) & "  " & c.name & "  Bottom=" & (c.Top + c.Height)

        End If
    End If

    ' コンテナは掘る（中身を見る）
    If isContainer Then
        WalkContainer c, maxBottom
    End If

    Next c

    
 


    
End Sub




Public Sub Fix_Page8_DailyLog_Once()

    'Debug.Print "[Fix_Page8] ENTER"



    Dim uf As Object: Set uf = frmEval
    Dim mp As MSForms.MultiPage: Set mp = uf.Controls("MultiPage1")
    Dim pg As MSForms.Page: Set pg = mp.Pages("Page8")

    Dim maxBottom As Double
    maxBottom = 0#
    
    pg.Controls("lstDailyLogList").Height = 140

    
    WalkContainer pg, maxBottom

    Dim needShrink As Double
    needShrink = Application.Max(0, maxBottom - mp.Height + 1)


    If needShrink > 0 Then
        pg.Controls("lstDailyLogList").Height = Application.Max(40, 140 - needShrink)

    Else
        pg.Controls("lstDailyLogList").Height = 140
    End If

    '検証ログ（結果だけ）
    maxBottom = 0#
    WalkContainer pg, maxBottom
    Static callN As Long: callN = callN + 1: Debug.Print "[Fix_Page8] call#" & callN & " needShrink=" & needShrink & "  NewBottom=" & maxBottom & "  Overflow=" & (maxBottom - mp.Height)

End Sub




Public Sub Fix_Page6_Walk_FrameScroll_Once()
    Dim f As Object
    Set f = frmEval.Controls("MultiPage1").Pages("Page6").Controls("Frame6")

    '表示枠をMP1に合わせる
    f.Height = frmEval.Controls("MultiPage1").Height
    f.ScrollBars = fmScrollBarsVertical
    f.ScrollTop = 0

    '中身の最大Bottom → ScrollHeight
    Dim maxBottom As Double: maxBottom = 0#
    Dim c As Object
    For Each c In f.Controls
        If c.Top + c.Height > maxBottom Then maxBottom = c.Top + c.Height
    Next c
    f.ScrollHeight = maxBottom + 12

#If APP_DEBUG Then
    Debug.Print "[Fix_Page6] H=" & f.Height & " ScrollH=" & f.ScrollHeight
#End If
End Sub




Public Sub Temp_SetScroll_Frame3_Page3()
    With frmEval.Controls("MultiPage1").Pages(2).Controls("Frame3")
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 602.35   ' 578.35 + 24
    End With
End Sub



Public Sub Temp_SetScroll_Frame7_Page7()
    With frmEval.Controls("MultiPage1").Pages(6).Controls("Frame7")
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 560.35 + 24
    End With
End Sub



Public Sub Temp_TestScroll_Frame7_Page7()
    With frmEval.Controls("MultiPage1").Pages(6).Controls("Frame7")
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 584.35
        MsgBox "ScrollBars=" & .ScrollBars & vbCrLf & "ScrollHeight=" & .ScrollHeight, vbInformation
    End With
End Sub



Public Sub Temp_SetScroll_Frame1_PostureTab()
    Dim mp As Object: Set mp = frmEval.Controls("MultiPage1")
    Dim pg As Object: Set pg = mp.Pages(mp.value)

    With pg.Controls("Frame1")
        .Height = mp.Height                '←表示枠に収める（ここが本丸）
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 396 + 24           '←中身Bottom(396) + 余白
    End With
End Sub





Public Sub Temp_SetScroll_Frame2_Page2()
    Dim mp As Object: Set mp = frmEval.Controls("MultiPage1")
    With mp.Pages(1).Controls("Frame2")
        .Height = mp.Height
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 464 + 24
    End With
End Sub

