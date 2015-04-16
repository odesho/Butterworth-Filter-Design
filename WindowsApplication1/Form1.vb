Public Class FilterDesign



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        TrueInvisible()


    End Sub



    Private Sub Enterbtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Enterbtn.Click



        Dim AttP As String = A_p.Text
        Dim AttS As String = A_s.Text
        Dim frqP As String = F_p.Text
        Dim frqS As String = F_s.Text
        Dim caps As String = capstxt.Text
        Dim ResA As String = restxt.Text



        Dim butt_n As Double = Math.Ceiling(buttr_order(AttP, AttS, frqP, frqS))
        Dim Res As Double = calcRes(caps, frqP, AttP, butt_n)

        ResBs(butt_n, ResA, Res, caps)

        Me.TabControl1.SelectedTab = TabPage2



    End Sub

    Private Function buttr_order(ByRef AttP, ByRef AttS, ByRef frqP, ByRef frqS) As Double

        Dim ans As Double = Nothing

        ans = (Math.Log10(Math.Sqrt((10 ^ (AttS / 10) - 1) / (10 ^ (AttP / 10) - 1)))) / (Math.Log10(frqS / frqP))


        Return ans
    End Function


    Private Function calcRes(ByRef caps, ByRef frqP, ByRef Attp, ByRef butt_n) As Double

        Dim ans As Double = Nothing

        Dim frqP2 = 10 ^ (Math.Log10(Math.Sqrt(10 ^ (Attp / 10) - 1)) / butt_n) * frqP

        ans = Math.Round(1 / (frqP2 * caps * 10 ^ -6), 0)

        Return ans
    End Function

    Private Sub ResBs(ByRef butt_n As Double, ByRef ResA As Double, ByRef Res As Double, ByRef caps As Double)


        Dim ResB1 As Double = Nothing
        Dim ResB2 As Double = Nothing
        Dim ResB3 As Double = Nothing
        Dim ResB4 As Double = Nothing


        If butt_n = 1 Then
            Label36.Visible = True
            Label35.Visible = False
            Label36.Text = "Order n = " & butt_n
            Label35.Text = ""

            TrueInvisible()

            Label33.Visible = True
            Label34.Visible = True

            ResB1 = ResA * (2 - 1.4142)

            PictureBox4.Visible = True

            Label33.Visible = True
            Label34.Visible = True

            Label33.Text = Res / 1000 & " kohm"
            Label34.Text = caps & " uF"





        ElseIf butt_n = 2 Then
            Label36.Visible = True
            Label35.Visible = False
            Label36.Text = "Order n = " & butt_n
            Label35.Text = ""

            TrueInvisible()

            ResB1 = ResA * (2 - 1.4142)

            Label7.Visible = True
            Label20.Visible = True
            Label21.Visible = True
            Label22.Visible = True
            Label23.Visible = True
            Label24.Visible = True


            PictureBox2.Visible = True

            Label7.Text = Res / 1000 & " kohm"
            Label20.Text = Res / 1000 & " kohm"

            Label21.Text = caps & " uF"
            Label22.Text = caps & " uF"

            Label23.Text = ResA / 1000 & " kohm"
            Label24.Text = ResB1 / 1000 & " kohm"




        ElseIf butt_n = 3 Then
            Label36.Visible = True
            Label35.Visible = False

            Label36.Text = "Order n = " & butt_n
            Label35.Text = ""

            TrueInvisible()

            ResB1 = ResA



            Label25.Visible = True
            Label26.Visible = True
            Label27.Visible = True
            Label28.Visible = True
            Label29.Visible = True
            Label30.Visible = True
            Label31.Visible = True
            Label32.Visible = True
            PictureBox3.Visible = True

            Label25.Text = Res / 1000 & " kohm"
            Label26.Text = Res / 1000 & " kohm"
            Label27.Text = Res / 1000 & " kohm"

            Label28.Text = caps & " uF"
            Label31.Text = caps & " uF"
            Label32.Text = caps & " uF"

            Label29.Text = ResA / 1000 & " kohm"
            Label30.Text = ResB1 / 1000 & " kohm"



        ElseIf butt_n = 4 Then
            Label36.Visible = True
            Label35.Visible = False
            TrueInvisible()

            Label36.Text = "Order n = " & butt_n
            Label35.Text = ""

            Label8.Visible = True
            Label9.Visible = True
            Label13.Visible = True
            Label14.Visible = True

            Label10.Visible = True
            Label15.Visible = True
            Label18.Visible = True
            Label19.Visible = True

            Label12.Visible = True
            Label16.Visible = True
            Label11.Visible = True
            Label17.Visible = True


            ResB1 = ResA * (2 - 0.7654)
            ResB2 = ResA * (2 - 1.8478)

            PictureBox1.Visible = True

            Label8.Text = Res / 1000 & " kohm"
            Label9.Text = Res / 1000 & " kohm"
            Label13.Text = Res / 1000 & " kohm"
            Label14.Text = Res / 1000 & " kohm"

            Label10.Text = caps & " uF"
            Label15.Text = caps & " uF"
            Label18.Text = caps & " uF"
            Label19.Text = caps & " uF"

            Label12.Text = ResA / 1000 & " kohm"
            Label16.Text = ResA / 1000 & " kohm"
            Label11.Text = ResB1 / 1000 & " kohm"
            Label17.Text = ResB2 / 1000 & " kohm"


        Else
            Label36.Visible = False
            Label35.Visible = True
            Label35.Text = "Order n = " & butt_n & "  Not Available"

        End If


    End Sub


    Private Sub TrueInvisible()

        PictureBox1.Visible = False
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = False

        Label8.Visible = False
        Label9.Visible = False
        Label13.Visible = False
        Label14.Visible = False

        Label10.Visible = False
        Label15.Visible = False
        Label18.Visible = False
        Label19.Visible = False

        Label12.Visible = False
        Label16.Visible = False
        Label11.Visible = False
        Label17.Visible = False

        Label7.Visible = False
        Label20.Visible = False
        Label21.Visible = False
        Label22.Visible = False
        Label23.Visible = False
        Label24.Visible = False

        Label25.Visible = False
        Label26.Visible = False
        Label27.Visible = False
        Label28.Visible = False
        Label29.Visible = False
        Label30.Visible = False
        Label31.Visible = False
        Label32.Visible = False

        Label33.Visible = False
        Label34.Visible = False


    End Sub
End Class
