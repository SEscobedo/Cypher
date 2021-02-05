Module Crypto
    Class configuración
        Property ArchivoInicio As String
        Property Alfabeto As String
        Property InicioLlave As Long
    End Class

    Public Const ALFABETO_STD = "aáâàäåbcçdeéêèëfghiíìîjklmnñoóôòðöpqrstuúûùv}wxyz ~ø÷AÁÂÀBCDEÉÊÈFGHIÍÌÎJKLMNÑOÓÒÔPQRSTUÚÙÛVWXYZ¬,.1234(567890[]{)!#$%&/=+*?¡¿-_:;""'|<>§"
    Public config As New configuración
    Public RANDOM_KEY As String = "clavebladbaldbalbjalbalkjbalkjblakjblkjdlvkjbalkbablablabjblakjbb"
    Public RANDOM_KEY_FILE As String = ""
    Public ALFABETO As String


    Function Encriptar(ByVal text As String, ByVal key As String, ByVal metodo As String) As String
        Dim res As String
        If metodo = "de reversión" Then 'método 1. Reversión de la cadena
            res = StrReverse(text)
        ElseIf metodo = "Perfect security" Then 'método 3. Perfect security
            res = PerfectSecurityMethod(text, RANDOM_KEY)
        ElseIf metodo = "Julio César" Then 'método 4. Cesar Method
            Dim k As Long = CLng(key)
            res = CesarMethod(text, k)
        ElseIf metodo = "Perfect security II" Then 'método 2. Perfect security II
            'generar nueva llave.

            Dim RANDOM_KEY_II = Mid$(RANDOM_KEY, config.InicioLlave, Len(RANDOM_KEY)) & Mid$(RANDOM_KEY, 1, config.InicioLlave - 1)
            Dim ic As Long = Val(config.InicioLlave)

            res = PerfectSecurityMethod(text, RANDOM_KEY_II)
            config.InicioLlave = ic + Len(res)
            If config.InicioLlave >= Len(RANDOM_KEY_II) Then config.InicioLlave = 1
            'insertar número de referencia:
            res = Trim(Str(ic)) & ";" & res
        Else
            res = Nothing
        End If
        Encriptar = res
    End Function

    Function DesEncriptar(ByVal text As String, ByVal key As String, ByVal metodo As String) As String
        Dim res As String
        '  Dim v() As String

        If metodo = "de reversión" Then 'método 1. Reversión de la cadena
            res = StrReverse(text)

        ElseIf metodo = "Perfect security" Then
            res = PerfectSecurityDecode(text, RANDOM_KEY)
        ElseIf metodo = "Perfect security II" Then 'método 3. Perfect security

            'separar el número
            Dim i As Long
            Dim c As String = ""
            Dim ic As Long

            For i = 1 To Len(text)
                c = Mid$(text, i, 1)
                If c = ";" Then
                    ic = Val(Mid$(text, 1, i))
                    text = Mid$(text, i + 1, Len(text))
                    Exit For
                ElseIf i > 100 Then
                    MsgBox("Error en el cyphertext. El código de incicio de clave está dañado", MsgBoxStyle.Exclamation)
                    DesEncriptar = text
                    Exit Function
                End If
            Next

            'generar la clave correcta

            If ic <= 0 Then
                MsgBox("Error en el cyphertext. El código de incicio de clave está dañado", MsgBoxStyle.Exclamation)
                DesEncriptar = text
                Exit Function
            ElseIf ic > Len(RANDOM_KEY) Then
                MsgBox("Error en el cyphertext. El código de incicio de clave es mayor que la longitud de la llave cargada", MsgBoxStyle.Exclamation)
                DesEncriptar = text
                Exit Function
            End If

            Dim keyII As String
            If ic = 1 Then
                keyII = RANDOM_KEY
            Else
                keyII = Mid$(RANDOM_KEY, ic, Len(RANDOM_KEY)) & Mid$(RANDOM_KEY, 1, ic - 1)
            End If

            'decodificar el texto
            res = PerfectSecurityMethod(text, keyII)


        ElseIf metodo = "Julio César" Then
            Dim k As Long = CLng(key)
            res = CesarMethodDecode(text, k)
        Else
            res = Nothing
        End If
        DesEncriptar = res
    End Function

    Function PerfectSecurityMethod(ByVal PT As String, ByVal k As String) As String
        '1. tomar cada letra y cambiarla por un alfabeto en función de la llave

        Dim i As Long
        Dim res As String = ""
        For i = 1 To Len(PT)
            res = res & CharOperation(Mid$(PT, i, 1), Mid$(k, i, 1))
        Next
        Return res
    End Function
    Function PerfectSecurityDecode(ByVal s As String, ByVal k As String) As String
        '1. tomar cada letra y cambiarla por un alfabeto en función de la llave

        Dim i As Long
        Dim res As String = ""
        For i = 1 To Len(s)
            res = res & CharOperation(Mid$(s, i, 1), Mid$(k, i, 1))
        Next
        Return res
    End Function

    Function CesarMethod(ByVal s As String, ByVal k As Integer) As String
        Dim i As Long
        Dim res As String = ""

        For i = 1 To Len(s)
            res = res & CharInterchange(Mid(s, i, 1), k, False)
        Next
        Return res
    End Function
    Function CesarMethodDecode(ByVal s As String, ByVal k As Integer) As String
        Dim i As Long
        Dim res As String = ""

        For i = 1 To Len(s)
            res = res & CharInterchange(Mid(s, i, 1), k, True)
        Next
        Return res
    End Function

    Function CharInterchange(ByVal c As String, ByVal k As Integer, ByVal inverse As Boolean) As String
        Dim res As String = ""
        Dim i, j As Long


        If inverse Then
            For i = 0 To Len(ALFABETO)
                If c = Mid$(ALFABETO, i + 1, 1) Then
                    If i - k < 0 Then
                        j = Len(ALFABETO) - ((k - i) Mod Len(ALFABETO)) + 1
                    Else
                        j = ((i - k) Mod Len(ALFABETO)) + 1
                    End If

                    res = Mid$(ALFABETO, j, 1)
                    Exit For
                End If
                res = c
            Next
        Else
            For i = 0 To Len(ALFABETO)
                If c = Mid$(ALFABETO, i + 1, 1) Then
                    j = ((i + k) Mod Len(ALFABETO)) + 1
                    res = Mid$(ALFABETO, j, 1)
                    Exit For
                End If
                res = c
            Next
        End If

        Return res
    End Function

    Function CharOperation(ByVal c1 As String, ByVal c2 As String) As String
        Dim res As String = ""
        Dim i, j, n As Long
        Dim flag1, flag2 As Boolean


        For i = 1 To Len(ALFABETO)
            If c1 = Mid$(ALFABETO, i, 1) Then
                flag1 = True
                For j = 1 To Len(ALFABETO)
                    If c2 = Mid$(ALFABETO, j, 1) Then
                        flag2 = True
                        If Not j = i Then
                            n = j - i
                            If n < 0 Then
                                n = Len(ALFABETO) + 1 + n
                            End If
                            res = Mid$(ALFABETO, n, 1)
                        Else
                            res = Mid$(ALFABETO, i, 1)
                        End If

                        Exit For
                    End If
                Next
            End If
            If flag2 Then Exit For
        Next

        If Not flag1 Then res = c1
        If Not flag2 Then res = c1
        Return res
    End Function


    Function AdaptarLlave(ByVal key As String) As String
        Dim i, j As Long
        Dim flag As Boolean = False
        Dim c, res As String
        res = ""
        Dim d = Len(ALFABETO)
        For i = 1 To d
            res = Replace$(res, Mid$(ALFABETO, i, 1), "")
        Next
        Return res
    End Function
End Module
