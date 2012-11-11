Option Strict On
Option Explicit On

<Microsoft.VisualBasic.ComClass()> Public Class EnhacedTextBox

    'VARIABLES
    Private cadena_validar As String = ""
    Private pres As requisito = requisito.Si
    Private color_err As Color = Color.Firebrick
    Private color_sel As Color = Color.Teal
    Private color_bien As Color = Color.Olive
    Private max_caracter As Integer = 0
    Private min_caracter As Integer = 0
    Private information As String = ""
    Private Met As modo = modo.NoValidar
    Private ancho_me As Integer = 126
    Private sep As String = ":"
    Private val_error As Boolean = False
    Private paso As Boolean = True
    Private Alt As Boolean
    Private _size As Size
    Private TextCambiado As Boolean = True
    Private _formato As String = "XX.XXX,XX"
    Private _formatear As Format = Format.NoFormatear
    Private DifLabel As Integer = 0

    'ENUMERADOS
    ' ENUMERADO DE MODO DE VALIDACION
    Enum modo
        NoValidar = 0
        Numerico = 1
        Texto = 2
        Alfanumerico = 3
        Simbolos = 4
        TextoEspacio = 5
        Minusculas = 6
        Mayusculas = 7
        Especificar = 8
    End Enum
    ' ENUMERADO DE VALIDACION NECESARIA
    Enum requisito
        No = 0
        Si = 1
    End Enum

    Enum Format
        Formatear = 0
        NoFormatear = 1
    End Enum
    '--FIN ENUMERADOS

    'PROPIEDADES

    'PROPIEDAD DE MODO DE VALIDACION
    ''' <summary>
    ''' Establece los modo de validacion
    ''' <paramref name="Numerico">Permite la insercion solo de numeros, ademas de retroceso, espacio y teclas de direccion</paramref>
    ''' <paramref name="Texto">Permite la insercion solo de Textos, ademas de retroceso, espacio y teclas de direccion</paramref>
    ''' <paramref name="Alfanumerico">Permite la insercion solo de valores Alfanumerico, espacio ademas de retroceso y teclas de direccion</paramref>
    ''' <paramref name="Simbolos">Permite la insercion solo de Simbolos, espacio ademas de retroceso y teclas de direccion</paramref>
    ''' <paramref name="TextoEspacio">Permite la insercion solo de TextoEspacio, espacio ademas de retroceso y teclas de direccion</paramref>    
    ''' <paramref name="TextoEspacio">Permite la insercion solo de TextoEspacio, espacio ademas de retroceso y teclas de direccion</paramref>    
    ''' <paramref name="Minusculas">Permite la insercion solo de Minusculas, espacio ademas de retroceso y teclas de direccion</paramref>    
    ''' <paramref name="Mayusculas">Permite la insercion solo de Mayusculas, espacio ademas de retroceso y teclas de direccion</paramref>    
    ''' <paramref name="Especificar">Permite especificar en la propiedad CadenaValida el valor a permitir en la caja de texto</paramref>    
    ''' <paramref name="NoValidar">No realiza ningun tipo de validacion</paramref>    
    ''' </summary>
    Property Validar() As modo
        Get
            Return Met
        End Get
        Set(ByVal value As modo)
            Met = value
        End Set
    End Property

    'PROPIEDAD COLOR DEL LABEL
    ''' <summary>
    ''' Estanblece el color del texto label
    ''' </summary>
    '''     ''' <remarks>
    ''' <example> This sample shows how to set the <c>ID</c> field.
    ''' <code>
    ''' Dim alice As New Employee
    ''' alice.ID = 1234
    ''' </code>
    ''' </example>
    ''' </remarks>
    Public Property ColorLabel() As Color
        Get
            Return label.ForeColor
        End Get
        Set(ByVal value As Color)
            label.ForeColor = value
            punto.ForeColor = value
        End Set
    End Property

    'PROPIEDAD FUENTE DEL LABEL
    ''' <summary>
    ''' Estanblece la fuente del label
    ''' </summary>
    Public Property FuenteLabel() As Font
        Get
            Return label.Font
        End Get
        Set(ByVal value As Font)
            label.Font = value
            punto.Font = value
        End Set
    End Property

    'SEPARADOR
    ''' <summary>
    ''' Estanblece el separador entre el label y texbox
    ''' </summary>
    Public Property Separador() As Char
        Get
            Return CChar(punto.Text)
        End Get
        Set(ByVal value As Char)
            punto.Text = value
            sep = value
        End Set
    End Property

    'IZQUIERDA A DERECHA
    ''' <summary>
    ''' Estanblece la alineacion del texto en el textbox
    ''' </summary>
    Public Property IzqdaADerecha() As System.Windows.Forms.RightToLeft
        Get
            Return txt.RightToLeft
        End Get
        Set(ByVal value As System.Windows.Forms.RightToLeft)
            txt.RightToLeft = value
            dato.RightToLeft = value
        End Set
    End Property

    'PROPIEDAD TEXTO DEL LABEL
    ''' <summary>
    ''' Estanblece el texto del label
    ''' </summary>
    Public Property TextoLabel() As String
        Get
            Return label.Text.ToString
        End Get
        Set(ByVal value As String)
            Dim f, temp As Integer

            TextCambiado = True

            DifLabel = label.Width

            If value = "" Then
                temp = label.Width + punto.Width
                label.Text = value
                f = label.Width - temp
                punto.Visible = False
                punto.Text = ""
                Me.CambiaTamañoTxt()
            Else
                punto.Text = sep
                If label.Text = "" Then
                    temp = label.Width - punto.Width
                    label.Text = value
                    f = label.Width - temp
                    punto.Visible = True
                Else
                    temp = label.Width
                    label.Text = value
                    f = label.Width - temp
                End If
                DifLabel = DifLabel - label.Width
                Me.CambiaTamañoTxt()
            End If


            punto.Location = New Point(label.Width, 4)
            txt.Location = New Point(label.Width + punto.Width + 2, 2)
            dato.Location = New Point(label.Width + punto.Width + 2, 5)

        End Set
    End Property

    'PROPIEDAD CURSOR LABEL
    ''' <summary>
    ''' Establece el cursor de label
    ''' </summary>
    Public Property CursorLabel() As Cursor
        Get
            Return label.Cursor
        End Get
        Set(ByVal value As Cursor)
            label.Cursor = value
        End Set
    End Property

    'PROPIEDAD DEL TEXTO DEL CAMPO
    ''' <summary>
    ''' Establece el texto inicial del textobox
    ''' </summary>
    Public Property TextoCampo() As String
        Get
            Return txt.Text.ToString
        End Get
        Set(ByVal value As String)
            txt.Text = value
            dato.Text = value
        End Set
    End Property

    'PROPIEDAD DEL COLOR DEL TEXTO DEL CAMPO
    ''' <summary>
    ''' Establece el color del texto del textbox
    ''' </summary>
    Public Property ColorTxt() As Color
        Get
            Return txt.ForeColor
        End Get
        Set(ByVal value As Color)
            txt.ForeColor = value
            Me.ColorCadena = value
        End Set
    End Property

    'PROPIEDAD DEL LA FUENTE DEL TEXTO DEL CAMPO
    ''' <summary>
    ''' Establece la fuente del textbox
    ''' </summary>
    Public Property FuenteTxt() As Font
        Get
            Return txt.Font
        End Get
        Set(ByVal value As Font)
            txt.Font = value
            Me.FuenteCadena = value
        End Set
    End Property

    'PROPIEDAD DEL CADENA A VALIDAR
    ''' <summary>
    ''' Especifica la cadena a validar dentro del textbox
    ''' </summary>
    Public Property CadenaValida() As String
        Get
            Return cadena_validar
        End Get
        Set(ByVal value As String)
            cadena_validar = value
        End Set
    End Property

    'PROPIEDAD VALIDACION NECESARIA
    ''' <summary>
    ''' Especifica si es obligatoria la validacion
    ''' </summary>
    Public Property Necesario() As requisito
        Get
            Return pres
        End Get
        Set(ByVal value As requisito)
            pres = value
        End Set
    End Property

    'PROPIEDADCOLOR VALIDACION ERRONEA
    ''' <summary>
    ''' Establece el color del texto label, ante un error de validacion
    ''' </summary>
    Public Property ColorError() As Color
        Get
            Return color_err
        End Get
        Set(ByVal value As Color)
            color_err = value
        End Set
    End Property

    'PROPIEDAD COLOR ON_FOCUS
    ''' <summary>
    ''' Establece el color del texto label, cuando el control obtiene el foco
    ''' </summary>
    Public Property ColorFoco() As Color
        Get
            Return color_sel
        End Get
        Set(ByVal value As Color)
            color_sel = value
        End Set
    End Property

    'PROPIEDAD COLOR VALIDACION SATISFACTORIA
    ''' <summary>
    ''' Establece el color del texto label, ante una validacion satisfactoria
    ''' </summary>
    Public Property ColorOk() As Color
        Get
            Return color_bien
        End Get
        Set(ByVal value As Color)
            color_bien = value
        End Set
    End Property

    'TAMAÑO

    ''' <summary>
    ''' Establece el tamaño del control
    ''' </summary>
    ''' 
    Public Overloads Property size() As Size
        Get
            Return New Size(Me.Width, Me.Height)
        End Get
        Set(ByVal value As Size)

            If paso Then
                _size = value
                paso = False
            Else
                _size = Me.size
            End If

            With Me
                .Width = value.Width
                .Height = value.Height
            End With

            If Me.Multilinea Then
                txt.Height = Me.Height - 2
            End If

            dato.Height = txt.Height

            'If TextCambiado Or (_size.Width < Me.Width) Then
            '    txt.Width = txt.Width + (Me.Width - _size.Width)
            '    dato.Width = txt.Width
            'End If
            Me.CambiaTamañoTxt()

            Me.Refresh()

            'TextCambiado = False

        End Set
    End Property

    'Public Property Tamaño() As Size
    '    Get
    '        Return txt.Size
    '    End Get
    '    Set(ByVal value As Size)
    '        txt.Size = value
    '    End Set
    'End Property


    'PROPIEDAD MAXIMO CARACTERES CAMPO
    ''' <summary>
    ''' Establece numero maximo de caracteres
    ''' </summary>
    Property NMaxCaracteres() As Integer
        Get
            Return txt.MaxLength
        End Get
        Set(ByVal value As Integer)
            txt.MaxLength = value
        End Set
    End Property

    'PROPIEDAD FONDO COLOR CAMPO
    ''' <summary>
    ''' Establece el color del cuadro de texto
    ''' </summary>
    Property FondoCampo() As Color
        Get
            Return txt.BackColor
        End Get
        Set(ByVal value As Color)
            txt.BackColor = value
        End Set
    End Property

    'PROPEDAD FUNDO COLOR CONTROL
    ''' <summary>
    ''' Establece el color de fondo del control
    ''' </summary>
    Public Property ColorFondo() As Color
        Get
            Return Me.BackColor
        End Get
        Set(ByVal value As Color)
            Me.BackColor = value
            dato.BackColor = value
        End Set
    End Property

    '

    'PROPIEDAD COLOR BORDE CAMPO
    ''' <summary>
    ''' Establece el borde del cuadro de texto
    ''' </summary>
    Property BordeCampo() As System.Windows.Forms.BorderStyle
        Get
            Return txt.BorderStyle
        End Get
        Set(ByVal value As System.Windows.Forms.BorderStyle)
            txt.BorderStyle = value
        End Set
    End Property

    'PROPIEDAD MINIMO DE CARACTERES REQUERIDOS
    ''' <summary>
    ''' Establece el minimo de caracteres permitido
    ''' </summary>
    Property NMinCaracteres() As Integer
        Get
            Return min_caracter
        End Get
        Set(ByVal value As Integer)
            min_caracter = value
        End Set
    End Property

    'PROPIEDAD ACEPTA RETORNO DE CARRO
    ''' <summary>
    ''' Establece si el control acepta el retorno de carro
    ''' </summary>
    Property AceptaReturn() As Boolean
        Get
            Return txt.AcceptsReturn
        End Get
        Set(ByVal value As Boolean)
            txt.AcceptsReturn = value
        End Set
    End Property

    'PROPIEDAD ACEPTA TABULACION
    ''' <summary>
    ''' Establece si el control acepta Tabulacion
    ''' </summary>
    Property AceptaTab() As Boolean
        Get
            Return txt.AcceptsTab
        End Get
        Set(ByVal value As Boolean)
            txt.AcceptsTab = value
        End Set
    End Property

    'PROPIEDAD CAMPO MILTILINEA
    ''' <summary>
    ''' Establece si el control es multilinea
    ''' </summary>
    Property Multilinea() As Boolean
        Get
            Return txt.Multiline
        End Get
        Set(ByVal value As Boolean)
            txt.Multiline = value
            If value = False Then
                Me.Height = 24
            End If
        End Set
    End Property

    'PROPIEDAD MAXIMO DE CARACTERES REQUERIDOS
    ''' <summary>
    ''' Establece numero maximo de caracteres requeridos
    ''' </summary>
    Property MaxCaracterRequerido() As Integer
        Get
            Return max_caracter
        End Get
        Set(ByVal value As Integer)
            max_caracter = value
        End Set
    End Property

    'PROPIEDAD TITULO GLOBO INFO
    ''' <summary>
    ''' Establece el titulos del globo de informacion
    ''' </summary>
    Property TituloInfo() As String
        Get
            Return Informacion.ToolTipTitle
        End Get
        Set(ByVal value As String)
            Informacion.ToolTipTitle = value
        End Set
    End Property

    'PROPIEDAD ICONO GLOBO INFO
    ''' <summary>
    ''' Establece el icono del globo de informacion
    ''' </summary>
    Property IconoInfo() As System.Windows.Forms.ToolTipIcon
        Get
            Return Informacion.ToolTipIcon
        End Get
        Set(ByVal value As System.Windows.Forms.ToolTipIcon)
            Informacion.ToolTipIcon = value
        End Set
    End Property

    'PROPIEDAD TEXTO GLOBO INFORMACION
    ''' <summary>
    ''' Establece el contenido del globo de informacion
    ''' </summary>
    Property Info() As String
        Get
            Return Informacion.GetToolTip(Me.label)
        End Get
        Set(ByVal value As String)
            Informacion.SetToolTip(Me.label, value)
            If value <> "" Then
                Me.CursorLabel = Cursors.Hand
            Else
                Me.CursorLabel = Cursors.Default
            End If
        End Set
    End Property

    'PROPIEDAD LETRA MAYUSCULA, MINUSCULA
    ''' <summary>
    ''' Establece como mostrar el texto, mayuscula o minuscula
    ''' </summary>
    Property MayusoMinus() As System.Windows.Forms.CharacterCasing
        Get
            Return txt.CharacterCasing
        End Get
        Set(ByVal value As System.Windows.Forms.CharacterCasing)
            txt.CharacterCasing = value
        End Set
    End Property

    'PROPIEDAD TEXTO DE TXTCADENA
    ''' <summary>
    ''' Establece el texto del label mascara mascara a mostrar inicialmente
    ''' </summary>
    Property TextoCadena() As String
        Get
            Return dato.Text
        End Get
        Set(ByVal value As String)
            dato.Text = value
        End Set
    End Property

    'PROPIEDAD COLOR DEL TXTCADENA
    ''' <summary>
    ''' Establece el color de label mascara
    ''' </summary>
    Property ColorCadena() As Color
        Get
            Return dato.ForeColor
        End Get
        Set(ByVal value As Color)
            dato.ForeColor = value
        End Set
    End Property

    'PROPIEDAD COLOR FUENTE TXTCADENA
    ''' <summary>
    ''' Establece la fuente del label mascara
    ''' </summary>
    Property FuenteCadena() As Font
        Get
            Return dato.Font
        End Get
        Set(ByVal value As Font)
            dato.Font = value
        End Set
    End Property

    'PROPIEDAD CURSORCADENA
    ''' <summary>
    ''' Establece el cursor para label mascara
    ''' </summary>
    Public Property CursorCadena() As Cursor
        Get
            Return dato.Cursor
        End Get
        Set(ByVal value As Cursor)
            dato.Cursor = value
        End Set
    End Property

    'PROPIEDAD CURSORCADENA
    ''' <summary>
    ''' Activa la funcion de formateo de texto
    ''' </summary>
    Public Property Formatear() As Format
        Get
            Return _formatear
        End Get
        Set(ByVal value As Format)
            If value = Format.NoFormatear Then
                _formatear = value
                Me.TextoCampo = ""
            Else
                _formatear = value
                Me.TextoCampo = "___________________________"
            End If

        End Set
    End Property

    'PROPIEDAD CURSORCADENA
    ''' <summary>
    ''' Establece el formato
    ''' </summary>
    Public Property Formato() As String
        Get
            Return _formato
        End Get
        Set(ByVal value As String)
            _formato = value
        End Set
    End Property

    '--FIN PROPIEDADES

    'COLORACION GOT FOCUS
    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt.GotFocus, TextBox1.GotFocus
        If Not val_error Then
            label.ForeColor = color_sel
            punto.ForeColor = color_sel
        End If
        txt.Select(0, 0)
        Alt = True
        '  val_error = True
    End Sub

    Private Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt.KeyDown, TextBox1.KeyDown
        Alt = e.Shift
    End Sub

    ' VALIDACION KEYPRESS
    Private Sub txt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt.KeyPress, TextBox1.KeyPress
        Select Case Met
            Case modo.Alfanumerico
                If Char.IsLetterOrDigit(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case modo.Mayusculas
                If Char.IsUpper(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case modo.Minusculas
                If Char.IsLower(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case modo.Numerico
                If Char.IsDigit(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case modo.Simbolos
                If Char.IsSymbol(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case modo.Texto
                If Char.IsLetter(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case modo.TextoEspacio
                If Char.IsLetter(e.KeyChar) Or Char.IsSeparator(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
                    e.Handled = False
                Else
                    e.Handled = True
                End If
            Case modo.NoValidar
                e.Handled = False
        End Select
    End Sub

    'EVENTO LOST FOCUS
    Private Sub txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt.LostFocus, TextBox1.LostFocus
        Dim i, x As Integer
        Dim valida As Boolean
        Dim cadena As String = cadena_validar
        Dim texto As String = txt.Text.ToString
        If texto = "" And pres = requisito.Si Then
            txt.Focus()
            label.ForeColor = color_err
            punto.ForeColor = color_err
            val_error = True
            Exit Sub
        Else
            If Met = modo.Especificar Then

                For i = 0 To texto.Length - 1
                    valida = False
                    For x = 0 To cadena.Length - 1
                        If texto(i) = cadena(x) Then
                            valida = True
                        End If
                        If cadena.Length - 1 = x And valida = False Then
                            txt.Text = texto
                            label.ForeColor = color_err
                            punto.ForeColor = color_err
                            val_error = True
                            Exit Sub
                        End If
                    Next
                Next
                txt.Text = texto
            End If
        End If
        label.ForeColor = color_bien
        punto.ForeColor = color_bien
        dato.Visible = True
        Me.TextoCadena = Me.TextoCampo
        InfoDato.SetToolTip(Me.dato, Me.TextoCampo)
        txt.Visible = False
        val_error = False
    End Sub

    Private Sub foco_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles foco.GotFocus
        If val_error Then
            SendKeys.Send(Chr(13) + vbTab)
            txt.Select(0, txt.Text.Length)
        Else
            SendKeys.Send(vbTab)
            Alt = False
        End If
    End Sub

    Private Sub foco1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles foco1.GotFocus
        If Alt Then
            SendKeys.Send(Chr(13) + vbTab)
            Alt = False
        Else
            txt.Visible = True
            dato.Visible = False
            txt.Focus()
        End If

    End Sub
    Private Sub dato_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dato.MouseClick
        txt.Visible = True
        dato.Visible = False
        txt.Focus()
    End Sub

    Private Sub formateardor()

    End Sub

    Private Sub CambiaTamañoTxt()
        If TextCambiado Or (_size.Width < Me.Width) Then
            txt.Width = ((Me.Width - (label.Width + punto.Width)) - 2)
            dato.Width = txt.Width
            TextCambiado = False
        End If
    End Sub

    Private Sub EnhacedTextBox_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.foco.Focus()
    End Sub
End Class