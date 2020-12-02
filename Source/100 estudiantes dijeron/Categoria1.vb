Imports System.Data.OleDb
Public Class frmJuego
    Dim mi_ContadorJ1 As Integer
    Dim mi_ContadorJ2 As Integer
    Dim mi_Puntos As Integer
    Dim mi_Contador As Integer
    Dim mi_Pregunta As Integer
    Dim rm_Numero As New Random
    Dim mi_Cont As Integer
    Dim mi_PosX As Integer
    Dim mi_PosY As Integer
    Dim ms_ruta As String
    Dim mi_ContImagen As Integer

    Dim num(64) As Integer

    'CONFIGURA INICIO DEL JUEGO
    Private Sub Carga_Juego()
        pbxJugador1.Load("C:\100 estudiantes dijeron\Images\UIN_logo.png")
        pbxJugador2.Load("C:\100 estudiantes dijeron\Images\UIN_logo.png")

        tbxPuntosJ1.Enabled = False
        tbxPuntosJ2.Enabled = False

        tbxPregunta1.Enabled = False
        tbxPregunta2.Enabled = False
        tbxPregunta3.Enabled = False
        tbxPregunta4.Enabled = False
        tbxPregunta5.Enabled = False

        tbxPuntos1.Enabled = False
        tbxPuntos2.Enabled = False
        tbxPuntos3.Enabled = False
        tbxPuntos4.Enabled = False
        tbxPuntos5.Enabled = False
        tbxTP.Enabled = False
        tbxPregunta.Enabled = False

        Me.cbxEquipo1.Items.Add("Fuerzas especiales ginyu")
        Me.cbxEquipo1.Items.Add("infobits")
        Me.cbxEquipo1.Items.Add("Ricos suaves")
        Me.cbxEquipo1.Items.Add("Los manda mas")
        Me.cbxEquipo2.Items.Add("Fuerzas especiales ginyu")
        Me.cbxEquipo2.Items.Add("infobits")
        Me.cbxEquipo2.Items.Add("Ricos suaves")
        Me.cbxEquipo2.Items.Add("Los manda mas")

        cbxEquipo1.SendToBack()
        cbxEquipo2.SendToBack()
        btnInicio.SendToBack()

        Call Sonido()
    End Sub

    'CARGA JUEGO
    Private Sub Carga_Categoria()
        Dim ds_Datos As New DataSet
        Dim dt_Datos As New DataTable
        Dim s_Sql As String
otra:
        mi_Pregunta = rm_Numero.Next(1, 46)

        If Evitar_Repeticion(mi_Pregunta) = True Then
            GoTo otra
        End If

        s_Sql = "SELECT Id_Juego, RES1, PUN1, RES2, PUN2, RES3, PUN3, RES4, PUN4, RES5, PUN5 FROM Juego WHERE Id_Juego=" & mi_Pregunta

        Dim da_datos As New OleDbDataAdapter(s_Sql, cn_Conexion)
        da_datos.Fill(dt_Datos)
        For Each DataRow In dt_Datos.Rows
            tbxPregunta.Text = DataRow(0)
            tbxPregunta1.Text = DataRow(1)
            tbxPuntos1.Text = DataRow(2)
            tbxPregunta2.Text = DataRow(3)
            tbxPuntos2.Text = DataRow(4)
            tbxPregunta3.Text = DataRow(5)
            tbxPuntos3.Text = DataRow(6)
            tbxPregunta4.Text = DataRow(7)
            tbxPuntos4.Text = DataRow(8)
            tbxPregunta5.Text = DataRow(9)
            tbxPuntos5.Text = DataRow(10)
        Next
    End Sub

    Private Sub Conf_Luces()
        pbxInicio.SendToBack()
        btnInicio.Enabled = True
        cbxEquipo1.Enabled = True
        cbxEquipo2.Enabled = True
        tmrImagen.Enabled = False
        ms_ruta = "C:\100 estudiantes dijeron\Images\"
        pbxInicio.Image = New Bitmap(ms_ruta + "luces.gif")

    End Sub

    Private Function Evitar_Repeticion(ByVal i_Pregunta) As Boolean
        Dim Cont As Integer
        Dim i_Conteo As Integer

        If mi_Contador = 0 Then
            num(0) = i_Pregunta
            Evitar_Repeticion = False
            mi_Contador = mi_Contador + 1
            Exit Function
        End If
        i_Conteo = mi_Contador

        For X = 0 To i_Conteo - 1
            Cont = 0

            If num(X) = i_Pregunta Then
                MsgBox("valor repetido")
                Evitar_Repeticion = True
                Cont = 1
                Exit Function
            End If
        Next
        If Cont = 0 Then
            num(mi_Contador) = i_Pregunta
            Evitar_Repeticion = False
            mi_Contador = mi_Contador + 1
        End If
    End Function

    Private Sub Ganar()
        Dim ruta As String
        ruta = My.Application.Info.DirectoryPath
        Dim s As System.Media.SoundPlayer
        s = New System.Media.SoundPlayer("C:\100 estudiantes dijeron\Binary\correcta.wav")
        s.Play()
    End Sub

    Private Sub Perder()
        Dim ruta As String
        ruta = My.Application.Info.DirectoryPath
        Dim s As System.Media.SoundPlayer
        s = New System.Media.SoundPlayer("C:\100 estudiantes dijeron\Binary\incorrecta.wav")
        s.Play()
    End Sub

    Private Sub Sonido()
        Dim ruta As String
        ruta = My.Application.Info.DirectoryPath
        Dim s As System.Media.SoundPlayer
        s = New System.Media.SoundPlayer("C:\100 estudiantes dijeron\Binary\inicio.wav")
        s.Play()
    End Sub

    'CONTROL DE TECLAS
    Private Sub Categoria1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            'DESTAPAR JUEGO CON SUMA DE PUNTOS
            Case Keys.Q
                Call Ganar()
                pbxP1.Visible = False
                pbxP11.Visible = False
                tbxTP.Text = Val(tbxTP.Text) + Val(tbxPuntos1.Text)
            Case Keys.W
                pbxP2.Visible = False
                Call Ganar()
                pbxP12.Visible = False
                tbxTP.Text = Val(tbxTP.Text) + Val(tbxPuntos2.Text)
            Case Keys.E
                Call Ganar()
                pbxP3.Visible = False
                pbxP13.Visible = False
                tbxTP.Text = Val(tbxTP.Text) + Val(tbxPuntos3.Text)
            Case Keys.R
                Call Ganar()
                pbxP4.Visible = False
                pbxP14.Visible = False
                tbxTP.Text = Val(tbxTP.Text) + Val(tbxPuntos4.Text)
            Case Keys.T
                Call Ganar()
                pbxP5.Visible = False
                pbxP15.Visible = False
                tbxTP.Text = Val(tbxTP.Text) + Val(tbxPuntos5.Text)
                'DESTAPAR SIN SUMA DE PUNTOS
            Case Keys.A
                Call Ganar()
                pbxP1.Visible = False
                pbxP11.Visible = False
            Case Keys.S
                Call Ganar()
                pbxP2.Visible = False
                pbxP12.Visible = False
            Case Keys.D
                Call Ganar()
                pbxP3.Visible = False
                pbxP13.Visible = False
            Case Keys.F
                Call Ganar()
                pbxP4.Visible = False
                pbxP14.Visible = False
            Case Keys.G
                Call Ganar()
                pbxP5.Visible = False
                pbxP15.Visible = False
                'DESTAPAR X JUGADOR 2
            Case Keys.L
                Call Perder()
                pbxX2.Load("C:\100 estudiantes dijeron\Images\x.png")
                pbxX2.Visible = True
            Case Keys.K
                Call Perder()
                pbxXX2.Load("C:\100 estudiantes dijeron\Images\x.png")
                pbxXX2.Visible = True
            Case Keys.J
                Call Perder()
                pbxXXX2.Load("C:\100 estudiantes dijeron\Images\x.png")
                pbxXXX2.Visible = True
                ' DESTAPAR X JUGADOR 1
            Case Keys.O
                Call Perder()
                pbxX1.Load("C:\100 estudiantes dijeron\Images\x.png")
                pbxX1.Visible = True
            Case Keys.I
                Call Perder()
                pbxXX1.Load("C:\100 estudiantes dijeron\Images\x.png")
                pbxXX1.Visible = True
            Case Keys.U
                Call Perder()
                pbxXXX1.Load("C:\100 estudiantes dijeron\Images\x.png")
                pbxXXX1.Visible = True
                'ASIGNAR PUNTOS
            Case Keys.Z
                tbxPuntosJ1.Text = Val(tbxPuntosJ1.Text) + Val(tbxTP.Text)
            Case Keys.X
                tbxPuntosJ2.Text = Val(tbxPuntosJ2.Text) + Val(tbxTP.Text)
                'SONIDO
            Case Keys.B
                Call Sonido()
                'QUITAR BANER
            Case Keys.M
                Call Conf_Luces()
                'RESETEAR
            Case Keys.N
                pbxP1.Visible = True
                pbxP11.Visible = True
                pbxP2.Visible = True
                pbxP12.Visible = True
                pbxP3.Visible = True
                pbxP13.Visible = True
                pbxP4.Visible = True
                pbxP14.Visible = True
                pbxP5.Visible = True
                pbxP15.Visible = True
                pbxX1.Visible = False
                pbxXX1.Visible = False
                pbxXXX1.Visible = False
                pbxX2.Visible = False
                pbxXX2.Visible = False
                pbxXXX2.Visible = False

                tbxTP.Text = 0
                Call Carga_Categoria()
            Case Keys.V
                cbxEquipo1.ResetText()
                cbxEquipo1.Enabled = True
                cbxEquipo2.ResetText()
                cbxEquipo2.Enabled = True
                btnInicio.Enabled = True
            Case Keys.C
                Me.Close()
        End Select
    End Sub

    'INICIO 
    Private Sub Categoria1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        mi_Contador = 0
        Call Carga_Juego()
        Call Carga_Categoria()
        tmrImagen.Enabled = True
        cbxEquipo1.Enabled = False
        cbxEquipo2.Enabled = False
        btnInicio.Enabled = False
    End Sub

    'PERSONAJES DE LOS EQUIPOS
    Private Sub tbxJugador1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pbxJugador1.Click

        mi_ContadorJ1 = mi_ContadorJ1 + 1

        Select Case mi_ContadorJ1
            Case 1
                pbxJugador1.Load("C:\100 estudiantes dijeron\Images\1.jpg")
            Case 2
                pbxJugador1.Load("C:\100 estudiantes dijeron\Images\2.jpg")
            Case 3
                pbxJugador1.Load("C:\100 estudiantes dijeron\Images\3.jpg")
            Case 4
                pbxJugador1.Load("C:\100 estudiantes dijeron\Images\4.jpg")
            Case 5
                pbxJugador1.Load("C:\100 estudiantes dijeron\Images\5.jpg")
            Case 6
                pbxJugador1.Load("C:\100 estudiantes dijeron\Images\6.jpg")
                mi_ContadorJ1 = 0
        End Select
    End Sub

    Private Sub tbxJugador2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pbxJugador2.Click
        mi_ContadorJ2 = mi_ContadorJ2 + 1

        Select Case mi_ContadorJ2
            Case 1
                pbxJugador2.Load("C:\100 estudiantes dijeron\Images\1.jpg")
            Case 2
                pbxJugador2.Load("C:\100 estudiantes dijeron\Images\2.jpg")
            Case 3
                pbxJugador2.Load("C:\100 estudiantes dijeron\Images\3.jpg")
            Case 4
                pbxJugador2.Load("C:\100 estudiantes dijeron\Images\4.jpg")
            Case 5
                pbxJugador2.Load("C:\100 estudiantes dijeron\Images\5.jpg")
            Case 6
                pbxJugador2.Load("C:\100 estudiantes dijeron\Images\6.jpg")
                mi_ContadorJ2 = 0
        End Select
    End Sub

    Private Sub tmrImagen_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrImagen.Tick
        mi_ContImagen = mi_ContImagen + 1
        Select Case mi_ContImagen
            Case 1
                pbxInicio.Load("C:\100 estudiantes dijeron\Images\1.png")
            Case 2
                pbxInicio.Load("C:\100 estudiantes dijeron\Images\2.png")
            Case 3
                pbxInicio.Load("C:\100 estudiantes dijeron\Images\3.png")
            Case 4
                pbxInicio.Load("C:\100 estudiantes dijeron\Images\4.png")
            Case Else
                mi_ContImagen = 0
                Exit Sub
        End Select
    End Sub

    Private Sub btnInicio_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInicio.Click
        cbxEquipo1.Enabled = False
        cbxEquipo2.Enabled = False
        btnInicio.Enabled = False
    End Sub
End Class