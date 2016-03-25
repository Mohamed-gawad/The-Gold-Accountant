
Public Class clsNumber
    Shared Function nTOword(x As String) As String
        Dim Ma, Mi, n, R, f
        Dim Right, B As String
        Dim y As String
        On Error GoTo A
        Ma = " جنيه مصري "
        Mi = " قرش "
        If CDbl(x) < 0 Then x = Math.Abs(Val(x))
        B = Format(CDbl(x), "000000000000.00")
        Right = Math.Abs(Val(B.Substring(B.Length - 2, 2)))
        n = Int(x)
        R = SHorof(n)

        f = SHorof(Right)

        y = ""

        If R <> "" And f <> "" Then y = R + Ma & " و " + f + Mi

        If R <> "" And f = "" Then y = R + Ma

        If R = "" And f <> "" Then y = f + Mi

        Return y + "فقط لا غير"
A:
    End Function
    Shared Function SHorof(x)
        Dim Letter1 As String = Nothing
        Dim Letter2 As String = Nothing
        Dim Letter3 As String = Nothing
        Dim Letter4 As String = Nothing
        Dim Letter5 As String = Nothing
        Dim Letter6 As String = Nothing
        Dim c, c1, c2, c3, c4, c5, c6, n

        n = Int(x)
        c = Format((n), "000000000000")
        c1 = Val(Mid(c, 12, 1))
        Select Case c1
            Case Is = 1 : Letter1 = "واحد"
            Case Is = 2 : Letter1 = "اثنان"
            Case Is = 3 : Letter1 = "ثلاثة"
            Case Is = 4 : Letter1 = "أربعة"
            Case Is = 5 : Letter1 = "خمسة"
            Case Is = 6 : Letter1 = "ستة"
            Case Is = 7 : Letter1 = "سبعة"
            Case Is = 8 : Letter1 = "ثمانية"
            Case Is = 9 : Letter1 = "تسعة"
        End Select

        c2 = Val(Mid(c, 11, 1))
        Select Case c2
            Case Is = 1 : Letter2 = "عشر"
            Case Is = 2 : Letter2 = "عشرون"
            Case Is = 3 : Letter2 = "ثلاثون"
            Case Is = 4 : Letter2 = "أربعون"
            Case Is = 5 : Letter2 = "خمسون"
            Case Is = 6 : Letter2 = "ستون"
            Case Is = 7 : Letter2 = "سبعون"
            Case Is = 8 : Letter2 = "ثمانون"
            Case Is = 9 : Letter2 = "تسعون"
        End Select

        If Letter1 <> "" And c2 > 1 Then Letter2 = Letter1 + " و" + Letter2
        If Letter2 = "" Then Letter2 = Letter1
        If c1 = 0 And c2 = 1 Then Letter2 = Letter2 + "ة"
        If c1 = 1 And c2 = 1 Then Letter2 = "إحدى عشر"
        If c1 = 2 And c2 = 1 Then Letter2 = "إثنى عشر"
        If c1 > 2 And c2 = 1 Then Letter2 = Letter1 + " " + Letter2
        c3 = Val(Mid(c, 10, 1))
        Select Case c3
            Case Is = 1 : Letter3 = "مائة"
            Case Is = 2 : Letter3 = "مئتان"
            Case Is > 2 : Letter3 = Left(SHorof(c3), Len(SHorof(c3)) - 1) + "مائة"
        End Select
        If Letter3 <> "" And Letter2 <> "" Then Letter3 = Letter3 + " و" + Letter2
        If Letter3 = "" Then Letter3 = Letter2

        c4 = Val(Mid(c, 7, 3))
        Select Case c4
            Case Is = 1 : Letter4 = "ألف"
            Case Is = 2 : Letter4 = "ألفان"
            Case 3 To 10 : Letter4 = SHorof(c4) + " آلاف"
            Case Is > 10 : Letter4 = SHorof(c4) + " ألف"
        End Select
        If Letter4 <> "" And Letter3 <> "" Then Letter4 = Letter4 + " و" + Letter3
        If Letter4 = "" Then Letter4 = Letter3
        c5 = Val(Mid(c, 4, 3))
        Select Case c5
            Case Is = 1 : Letter5 = "مليون"
            Case Is = 2 : Letter5 = "مليونان"
            Case 3 To 10 : Letter5 = SHorof(c5) + " ملايين"
            Case Is > 10 : Letter5 = SHorof(c5) + " مليون"
        End Select
        If Letter5 <> "" And Letter4 <> "" Then Letter5 = Letter5 + " و" + Letter4
        If Letter5 = "" Then Letter5 = Letter4

        c6 = Val(Mid(c, 1, 3))
        Select Case c6
            Case Is = 1 : Letter6 = "مليار"
            Case Is = 2 : Letter6 = "ملياران"
            Case Is > 2 : Letter6 = SHorof(c6) + " مليار"
        End Select
        If Letter6 <> "" And Letter5 <> "" Then Letter6 = Letter6 + " و" + Letter5
        If Letter6 = "" Then Letter6 = Letter5
        Return Letter6

    End Function
End Class
