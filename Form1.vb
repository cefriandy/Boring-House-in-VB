Imports System.IO

Public Class Form1
    Dim graphics As Graphics
    Dim bmp As Bitmap
    Dim vertex(9), vertex2(9), vr(9), vr2(9), vr3(9), vr4(9), vr5(9), vr6(9), vs(9) As Tpoint
    Dim line(14), line2(11) As Tline
    Dim Trans(3, 3), Screen(3, 3), T12(3, 3), T3(3, 3), T4(3, 3), T5(3, 3), T6(3, 3), T7(3, 3) As Single
    Dim fp, bp As Single
    Dim window, dot, dot2, dot3 As New List(Of Point)
    Dim vrp, vpn, vup, cop, N, up, upa, V, U, cw, dop, r As Tpoint
    Dim wmax, wmin As Point
    Dim lvpn, lvup, lupa As Single
    Dim cube As Boolean = False
    Dim pen As Pen

    Structure Tpoint
        Dim x, y, z, w As Single
    End Structure

    Structure Tline
        Dim p1, p2 As Integer
    End Structure

    Sub setPoint(ByRef p As Tpoint, ByVal x As Single, ByVal y As Single, ByVal z As Single)
        p.x = x
        p.y = y
        p.z = z
        p.w = 1
    End Sub

    Sub setline(ByRef L As Tline, ByVal n1 As Integer, ByVal n2 As Integer)
        L.p1 = n1
        L.p2 = n2
    End Sub

    Sub DrawWindow()
        window.Add(New Point(50, 50))
        window.Add(New Point(250, 50))
        window.Add(New Point(250, 250))
        window.Add(New Point(50, 250))

        graphics.DrawLine(Pens.Black, window(0).X, window(0).Y, window(1).X, window(1).Y)
        graphics.DrawLine(Pens.Black, window(1).X, window(1).Y, window(2).X, window(2).Y)
        graphics.DrawLine(Pens.Black, window(2).X, window(2).Y, window(3).X, window(3).Y)
        graphics.DrawLine(Pens.Black, window(3).X, window(3).Y, window(0).X, window(0).Y)

        PictureBox1.Image = bmp
    End Sub

    Sub setpointline()
        setPoint(vertex(0), -1, -1, 1)
        setPoint(vertex(1), 1, -1, 1)
        setPoint(vertex(2), 1, 0, 1)
        setPoint(vertex(3), 0, 1, 1)
        setPoint(vertex(4), -1, 0, 1)
        setPoint(vertex(5), -1, -1, -1)
        setPoint(vertex(6), 1, -1, -1)
        setPoint(vertex(7), 1, 0, -1)
        setPoint(vertex(8), 0, 1, -1)
        setPoint(vertex(9), -1, 0, -1)

        setline(line(0), 0, 1)
        setline(line(1), 1, 2)
        setline(line(2), 2, 3)
        setline(line(3), 3, 4)
        setline(line(4), 4, 0)

        setline(line(5), 5, 6)
        setline(line(6), 6, 7)
        setline(line(7), 7, 8)
        setline(line(8), 8, 9)
        setline(line(9), 9, 5)
        setline(line(10), 0, 5)
        setline(line(11), 1, 6)
        setline(line(12), 2, 7)
        setline(line(13), 3, 8)
        setline(line(14), 4, 9)
    End Sub

    Sub setMatrix(ByRef M(,) As Single, ByVal col As Integer, ByVal a As Single, ByVal b As Single, ByVal c As Single, ByVal d As Single)
        M(0, col) = a
        M(1, col) = b
        M(2, col) = c
        M(3, col) = d
    End Sub

    Function MultiplyMat(p As Tpoint, M(,) As Single) As Tpoint
        Dim result As Tpoint
        Dim w As Single
        w = (p.x * M(0, 3) + p.y * M(1, 3) + p.z * M(2, 3) + p.w * M(3, 3))
        result.x = (p.x * M(0, 0) + p.y * M(1, 0) + p.z * M(2, 0) + p.w * M(3, 0)) / w
        result.y = (p.x * M(0, 1) + p.y * M(1, 1) + p.z * M(2, 1) + p.w * M(3, 1)) / w
        result.z = (p.x * M(0, 2) + p.y * M(1, 2) + p.z * M(2, 2) + p.w * M(3, 2)) / w
        result.w = 1
        Return result
    End Function

    Sub calculate()
        'N
        vpn.x = vpnXbox.Text
        vpn.y = vpnYbox.Text
        vpn.z = vpnZbox.Text

        lvpn = Math.Sqrt((vpn.x * vpn.x) + (vpn.y * vpn.y) + (vpn.z * vpn.z))

        N.x = (vpn.x / lvpn)
        N.y = (vpn.y / lvpn)
        N.z = (vpn.z / lvpn)

        'up
        vup.x = vupXbox.Text
        vup.y = vupYbox.Text
        vup.z = vupZbox.Text

        lvup = Math.Sqrt((vup.x * vup.x) + (vup.y * vup.y) + (vup.z * vup.z))

        up.x = (vup.x / lvup)
        up.y = (vup.y / lvup)
        up.z = (vup.z / lvup)

        ' up'
        upa.x = (up.x - (((up.x * N.x) + (up.y * N.y) + (up.z * N.z)) * N.x))
        upa.y = (up.y - (((up.x * N.x) + (up.y * N.y) + (up.z * N.z)) * N.y))
        upa.z = (up.z - (((up.x * N.x) + (up.y * N.y) + (up.z * N.z)) * N.z))

        lupa = Math.Sqrt((upa.x * upa.x) + (upa.y * upa.y) + (upa.z * upa.z))

        'V
        V.x = upa.x / lupa
        V.y = upa.y / lupa
        V.z = upa.z / lupa

        'U = V x N
        U.x = ((V.y * N.z) - (V.z * N.y))
        U.y = ((V.z * N.x) - (V.x * N.z))
        U.z = ((V.x * N.y) - (V.y * N.x))

        fp = FPbox.Text
        bp = BPbox.Text

        'rumus CW  
        wmax.X = WmaxXbox.Text
        wmax.Y = WmaxYbox.Text
        wmin.X = WminXbox.Text
        wmin.Y = WminYbox.Text

        cw.x = ((wmax.X + wmin.X) / 2)
        cw.y = ((wmax.Y + wmin.Y) / 2)
        cw.z = 0

        'Rumus DOP 
        cop.x = copXbox.Text
        cop.y = copYbox.Text
        cop.z = copZbox.Text

        dop.x = (cw.x - cop.x)
        dop.y = (cw.y - cop.y)
        dop.z = (cw.z - cop.z)

        'rumus r  
        vrp.x = vrpXbox.Text
        vrp.y = vrpYbox.Text
        vrp.z = vrpZbox.Text

        'T1T2  
        setMatrix(T12, 0, U.x, U.y, U.z, (-vrp.x * U.x - vrp.y * U.y - vrp.z * U.z))
        setMatrix(T12, 1, V.x, V.y, V.z, (-vrp.x * V.x - vrp.y * V.y - vrp.z * V.z))
        setMatrix(T12, 2, N.x, N.y, N.z, (-vrp.x * N.x - vrp.y * N.y - vrp.z * N.z))
        setMatrix(T12, 3, 0, 0, 0, 1)

        'T3
        setMatrix(T3, 0, 1, 0, 0, -cop.x)
        setMatrix(T3, 1, 0, 1, 0, -cop.y)
        setMatrix(T3, 2, 0, 0, 1, -cop.z)
        setMatrix(T3, 3, 0, 0, 0, 1)

        'T4
        setMatrix(T4, 0, 1, 0, -dop.x / dop.z, 0)
        setMatrix(T4, 1, 0, 1, -dop.y / dop.z, 0)
        setMatrix(T4, 2, 0, 0, 1, 0)
        setMatrix(T4, 3, 0, 0, 0, 1)

        ' T5
        setMatrix(T5, 0, 1, 0, 0, 0)
        setMatrix(T5, 1, 0, 1, 0, 0)
        setMatrix(T5, 2, 0, 0, 1, cop.z)
        setMatrix(T5, 3, 0, 0, 0, 1)

        'T6
        setMatrix(T6, 0, 2 / (wmax.X - wmin.X), 0, 0, 0)
        setMatrix(T6, 1, 0, 2 / (wmax.Y - wmin.Y), 0, 0)
        setMatrix(T6, 2, 0, 0, 1, 0)
        setMatrix(T6, 3, 0, 0, 0, 1)

        'T7
        setMatrix(T7, 0, 1, 0, 0, 0)
        setMatrix(T7, 1, 0, 1, 0, 0)
        setMatrix(T7, 2, 0, 0, 1, 0)
        setMatrix(T7, 3, 0, 0, -1 / cop.z, 1)

        setMatrix(Screen, 0, 100, 0, 0, 150)
        setMatrix(Screen, 1, 0, -100, 0, 150)
        setMatrix(Screen, 2, 0, 0, 0, 0)
        setMatrix(Screen, 3, 0, 0, 0, 1)

        If cube = True Then
            For i = 0 To vertex2.Count - 1
                vr(i) = MultiplyMat(vertex2(i), T12)
                vr2(i) = MultiplyMat(vr(i), T3)
                vr3(i) = MultiplyMat(vr2(i), T4)
                vr4(i) = MultiplyMat(vr3(i), T5)
                vr5(i) = MultiplyMat(vr4(i), T6)
                vr6(i) = MultiplyMat(vr5(i), T7)
                vs(i) = MultiplyMat(vr6(i), Screen)
            Next
        Else
            For i = 0 To vertex.Count - 1
                vr(i) = MultiplyMat(vertex(i), T12)
                vr2(i) = MultiplyMat(vr(i), T3)
                vr3(i) = MultiplyMat(vr2(i), T4)
                vr4(i) = MultiplyMat(vr3(i), T5)
                vr5(i) = MultiplyMat(vr4(i), T6)
                vr6(i) = MultiplyMat(vr5(i), T7)
                vs(i) = MultiplyMat(vr6(i), Screen)
            Next
        End If
        DrawWindow()
        clip()
    End Sub

    Sub clip()
        DrawWindow()
        Dim A, B As Point
        Dim C, in1, in2 As Point
        Dim pdx, pdy, index, i As Integer
        Dim Tex, Tey, Tlx, tly As Double
        Dim TEnter, TLeaving, T As Double
        Dim status1, status2, cw As Boolean

        If cube = True Then
            index = 11
        Else
            index = 14
        End If

        While index > -1
            cw = True
            If cube = True Then
                If index < 4 Then
                    pen = New Pen(Drawing.Color.Red, 1)
                Else
                    pen = New Pen(Drawing.Color.Black, 1)
                End If

                A.X = vs(line2(index).p1).x
                A.Y = vs(line2(index).p1).y
                B.X = vs(line2(index).p2).x
                B.Y = vs(line2(index).p2).y
            Else
                If index < 5 Then
                    pen = New Pen(Drawing.Color.Red, 1)
                Else
                    pen = New Pen(Drawing.Color.Black, 1)
                End If

                A.X = vs(line(index).p1).x
                A.Y = vs(line(index).p1).y
                B.X = vs(line(index).p2).x
                B.Y = vs(line(index).p2).y
            End If

            i = 0
            TLeaving = 1
            TEnter = 0
            While i < window.Count
                C = window(i)
                If cw = False Then
                    pdy = window((i + 1) Mod window.Count).Y - C.Y
                    pdx = C.X - window((i + 1) Mod window.Count).X
                Else
                    pdy = C.Y - window((i + 1) Mod window.Count).Y
                    pdx = window((i + 1) Mod window.Count).X - C.X
                End If
                status1 = (A.X - C.X) * pdy + (A.Y - C.Y) * pdx >= 0
                status2 = (B.X - C.X) * pdy + (B.Y - C.Y) * pdx >= 0

                If (status1 = False And status2 = False) Then
                    i = window.Count

                Else
                    If status1 = True And status2 = True Then

                    ElseIf status1 = False And status2 = True Then
                        T = ((A.X - C.X) * pdy + (A.Y - C.Y) * pdx) / ((A.X - B.X) * pdy + (A.Y - B.Y) * pdx)

                        If T > TEnter Then
                            TEnter = T

                        End If
                    ElseIf status1 = True And status2 = False Then
                        T = ((A.X - C.X) * pdy + (A.Y - C.Y) * pdx) / ((A.X - B.X) * pdy + (A.Y - B.Y) * pdx)
                        If T < TLeaving Then
                            TLeaving = T

                        End If

                    End If
                End If
                i = i + 1
            End While
            If status1 = False And status2 = False Then
            ElseIf TEnter > TLeaving Then
            Else

                If TEnter > 0 And TLeaving >= 1 Then
                    Tex = Math.Round(A.X + TEnter * (B.X - A.X))
                    Tey = Math.Round(A.Y + TEnter * (B.Y - A.Y))
                    in1.X = Tex
                    in1.Y = Tey
                    graphics.DrawLine(pen, B, in1)
                    graphics.DrawLine(pen, B, in1)
                End If

                If TEnter <= 0 And TLeaving < 1 Then
                    Tlx = A.X + TLeaving * (B.X - A.X)
                    tly = A.Y + TLeaving * (B.Y - A.Y)
                    in2.X = Tlx
                    in2.Y = tly
                    graphics.DrawLine(pen, A, in2)
                    graphics.DrawLine(pen, A, in2)
                End If

                If TLeaving < 1 And TEnter > 0 Then
                    Tlx = A.X + TLeaving * (B.X - A.X)
                    tly = A.Y + TLeaving * (B.Y - A.Y)
                    Tex = A.X + TEnter * (B.X - A.X)
                    Tey = A.Y + TEnter * (B.Y - A.Y)
                    in1.X = Tex
                    in1.Y = Tey
                    in2.X = Tlx
                    in2.Y = tly
                    graphics.DrawLine(pen, in1, in2)
                    graphics.DrawLine(pen, in1, in2)
                End If

                If (TLeaving >= 1 And TEnter <= 0) Then
                    graphics.DrawLine(pen, A, B)
                    graphics.DrawLine(pen, A, B)
                End If
            End If
            index = index - 1
        End While

    End Sub

    'Sub Drawhouse()
    '    Dim i, j As Integer
    '    If cube = True Then
    '        For i = 0 To line2.Count - 1
    '            graphics.DrawLine(Pens.Black, vs(line2(i).p1).x, vs(line2(i).p1).y, vs(line2(i).p2).x, vs(line2(i).p2).y)
    '        Next
    '        For j = 0 To 3
    '            graphics.DrawLine(Pens.Red, vs(line2(j).p1).x, vs(line2(j).p1).y, vs(line2(j).p2).x, vs(line2(j).p2).y)
    '        Next
    '    Else
    '        For i = 5 To 25
    '            graphics.DrawLine(Pens.Black, vs(line(i).p1).x, vs(line(i).p1).y, vs(line(i).p2).x, vs(line(i).p2).y)
    '        Next
    '        For j = 0 To 4
    '            graphics.DrawLine(Pens.Red, vs(line(j).p1).x, vs(line(j).p1).y, vs(line(j).p2).x, vs(line(j).p2).y)
    '        Next
    '    End If
    '    PictureBox1.Image = bmp
    'End Sub

    Private Sub View_Click(sender As Object, e As EventArgs) Handles View.Click
        bmp = New Bitmap(PictureBox1.Width, PictureBox1.Height)
        graphics = Graphics.FromImage(bmp)
        graphics.Clear(Color.White)
        setpointline()
        calculate()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bmp = New Bitmap(PictureBox1.Width, PictureBox1.Height)
        graphics = Graphics.FromImage(bmp)
        Button1.Enabled = False
        setpointline()
        calculate()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        OpenFileDialog1.Filter = "Text Files | *.txt"
        OpenFileDialog1.DefaultExt = "txt"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Dim pathfile As String = OpenFileDialog1.FileName
            Dim SW As New StreamReader(pathfile)
            Dim var As String
            Dim mode, phase, X, Y, Z, i, j As Single
            Dim p1, p2 As Integer

            mode = 0
            phase = 0
            j = 0
            i = 0
            X = 0
            Y = 0
            Z = 0

            var = CStr(SW.ReadLine)
            While var <> "End"
                If mode = 0 Then
                    If var = "VER" Then
                        mode = 1
                    ElseIf var = "LINE" Then
                        mode = 2
                    End If
                End If
                If mode = 1 Then
                    Select Case phase
                        Case 0
                            phase = 1
                        Case 1
                            X = (CSng(var))
                            phase = 2
                        Case 2
                            Y = (CSng(var))
                            phase = 3
                        Case 3
                            Z = (CSng(var))
                            setPoint(vertex2(i), X, Y, Z)
                            i = i + 1
                            X = 0
                            Y = 0
                            Z = 0
                            phase = 0
                            mode = 0
                    End Select
                ElseIf mode = 2 Then
                    Select Case phase
                        Case 0
                            phase = 1
                        Case 1
                            p1 = (CInt(var))
                            phase = 2
                        Case 2
                            p2 = (CInt(var))
                            setline(line2(j), p1, p2)
                            j = j + 1
                            p1 = 0
                            p2 = 0
                            phase = 0
                            mode = 0
                    End Select

                End If
                var = CStr(SW.ReadLine)

            End While
            SW.Close()

            bmp = New Bitmap(PictureBox1.Width, PictureBox1.Height)
            graphics = Graphics.FromImage(bmp)
            graphics.Clear(Color.White)
            cube = True
            calculate()
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        graphics.Clear(Color.White)
        cube = False
        setpointline()
        calculate()
        Button1.Enabled = False
    End Sub
End Class
