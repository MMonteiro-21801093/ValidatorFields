Imports System.Text.RegularExpressions
Imports System.Text

Public Class ValidatorFields
    Public Function isValidIban(ByVal iban As String) As Boolean 'ok
        If iban.Trim = "" Then
            Return True
        End If
        If iban.Trim.Length < 21 Then
            Return False
        End If

        iban = iban.ToUpper
        If iban.Length = 21 Then
            iban = "PT50" & iban
        End If



        iban = iban.Substring(4) & iban.Substring(0, 4)
        iban = iban.Replace("A", "10")
        iban = iban.Replace("B", "11")
        iban = iban.Replace("C", "12")
        iban = iban.Replace("D", "13")
        iban = iban.Replace("E", "14")
        iban = iban.Replace("F", "15")
        iban = iban.Replace("G", "16")
        iban = iban.Replace("H", "17")
        iban = iban.Replace("I", "18")
        iban = iban.Replace("J", "19")
        iban = iban.Replace("K", "20")
        iban = iban.Replace("L", "21")
        iban = iban.Replace("M", "22")
        iban = iban.Replace("N", "23")
        iban = iban.Replace("O", "24")
        iban = iban.Replace("P", "25")
        iban = iban.Replace("Q", "26")
        iban = iban.Replace("R", "27")
        iban = iban.Replace("S", "28")
        iban = iban.Replace("T", "29")
        iban = iban.Replace("U", "30")
        iban = iban.Replace("V", "31")
        iban = iban.Replace("W", "32")
        iban = iban.Replace("X", "33")
        iban = iban.Replace("Y", "34")
        iban = iban.Replace("Z", "35")

        Dim wl As Decimal = 0

        Try
            wl = CType(iban, Decimal) Mod 97
        Catch ex As Exception
            Return False
        End Try

        If wl <> 1 Then
            Return False
        End If

        Return True
    End Function
    Public Function isValidNif(ByVal nif As String) As Boolean 'ok
        If nif.Trim = "" Then
            Return True
        End If

        If (String.IsNullOrWhiteSpace(nif) _
                 OrElse (Not Regex.IsMatch(nif, "^[0-9]+$") _
                 OrElse (nif.Length <> 9))) Then
            Return False
        End If
        Dim c As Char = nif(0)

        Dim checkDigit As Integer = (Convert.ToInt32(c.ToString) * 9)
        Dim i As Integer = 2
        Do While (i <= 8)
            checkDigit = (checkDigit + (Convert.ToInt32(nif((i - 1)).ToString) * (10 - i)))
            i = (i + 1)
        Loop
        checkDigit = (11 - (checkDigit Mod 11))

        If (checkDigit >= 10) Then
            checkDigit = 0
        End If

        If (checkDigit.ToString <> nif(8).ToString) Then
            Return False
        End If
        Return True
    End Function
    Public Function isValidEmail(ByVal cp As String) As Boolean 'ok
        If cp.Trim = "" Then
            Return True
        End If
        'cria o padrão regex
        Dim padraoRegex As String = "^[-a-zA-Z0-9-_][-.a-zA-Z0-9-_]*@[-.a-zA-Z0-9]+(\.[-.a-zA-Z0-9]+)*\.(com|edu|info|gov|int|mil|net|org|biz|name|museum|coop|aero|pro|tv|pt|[a-zA-Z]{2})$"
        Dim verifica As New RegularExpressions.Regex(padraoRegex)
        'variavel boolean para tratar o status
        Dim valida As Boolean = False
        'verifica se o recurso foi fornecido
        If cp = "" Then
            'cep invalido
            valida = True
        Else
            'usa o método IsMatch Method para validar o regex
            valida = verifica.IsMatch(cp, 0)
        End If
        Return valida
    End Function
    Public Function isValidUrl(ByVal cp As String) As Boolean 'ok
        If cp.Trim = "" Then
            Return True
        End If
        'cria o padrão regex
        Dim padraoRegex As String = "^((http)|(https)|(ftp)):\/\/([\- \w]+\.)+\w{2,3}(\/ [%\-\w]+(\.\w{2,})?)*$"
        Dim verifica As New RegularExpressions.Regex(padraoRegex)
        'variavel boolean para tratar o status
        Dim valida As Boolean = False
        'verifica se o recurso foi fornecido
        If cp = "" Then
            'cep invalido
            valida = True
        Else
            'usa o método IsMatch Method para validar o regex
            valida = verifica.IsMatch(cp, 0)
        End If
        Return valida
    End Function
    Public Function isValidSoMinusculas(ByVal cp As String) As String ' ok
        If cp.Trim = "" Then
            Return cp
        End If

        Return cp.ToLower
    End Function
    Public Function isValidSoMaiuscSemEspeciais(ByVal cp As String) As String ' ok
        If cp.Trim = "" Then
            Return cp
        End If
        Dim codifica As String = cp.ToUpper

        codifica = codifica.Replace("Ã", "A")
        codifica = codifica.Replace("Á", "A")
        codifica = codifica.Replace("À", "A")
        codifica = codifica.Replace("Â", "A")

        codifica = codifica.Replace("É", "E")
        codifica = codifica.Replace("È", "E")
        codifica = codifica.Replace("Ê", "E")

        codifica = codifica.Replace("Í", "I")
        codifica = codifica.Replace("Ì", "I")

        codifica = codifica.Replace("Õ", "O")
        codifica = codifica.Replace("Ó", "O")
        codifica = codifica.Replace("Ò", "O")
        codifica = codifica.Replace("Ô", "O")

        codifica = codifica.Replace("Ú", "U")
        codifica = codifica.Replace("Ù", "U")
        codifica = codifica.Replace("Û", "U")

        codifica = codifica.Replace("Ç", "C")

        Return codifica
    End Function
    Public Function isValidNome(ByVal cp As String) As String ' ok
        If cp.Trim = "" Then
            Return cp
        End If
        Dim codifica As String = cp.ToUpper & " "

        codifica = codifica.Replace(" , ", " ")
        codifica = codifica.Replace(" ; ", " ")
        codifica = codifica.Replace(" & ", " ")
        codifica = codifica.Replace(" - ", " ")
        codifica = codifica.Replace(" . ", " ")
        codifica = codifica.Replace(" _ ", " ")


        codifica = codifica.Replace(" º ", " ")
        codifica = codifica.Replace(" ª ", " ")

        codifica = codifica.Replace(" DO ", " ")
        codifica = codifica.Replace(" DOS ", " ")
        codifica = codifica.Replace(" DA ", " ")
        codifica = codifica.Replace(" DAS ", " ")
        codifica = codifica.Replace(" DE ", " ")
        codifica = codifica.Replace(" DES ", " ")
        codifica = codifica.Replace(" D' ", " ")

        codifica = codifica.Replace(" A ", " ")
        codifica = codifica.Replace(" E ", " ")
        codifica = codifica.Replace(" I ", " ")
        codifica = codifica.Replace(" O ", " ")
        codifica = codifica.Replace(" U ", " ")

        codifica = codifica.Replace(" SA ", " ")
        codifica = codifica.Replace(" S.A ", " ")
        codifica = codifica.Replace(" S.A. ", " ")
        codifica = codifica.Replace(" S A. ", " ")
        codifica = codifica.Replace(" S A ", " ")
        codifica = codifica.Replace(" SA. ", " ")
        codifica = codifica.Replace(" S. A. ", " ")

        codifica = codifica.Replace(" LDA ", " ")
        codifica = codifica.Replace(" LDA.", " ")
        codifica = codifica.Replace(" LD.ª ", " ")
        codifica = codifica.Replace(" LDª. ", " ")
        codifica = codifica.Replace(" LDª ", " ")
        codifica = codifica.Replace(" LD ", " ")
        codifica = codifica.Replace(" L D A ", " ")
        codifica = codifica.Replace(" L.D.A. ", " ")
        codifica = codifica.Replace(" L. D. A. ", " ")

        Return codifica.Trim
    End Function
    Public Function isValidPrimeira(ByVal cp As String) As String ' ok
        If cp.Trim = "" Then
            Return cp
        End If
        Dim codifica As String = cp.Split(" ")(0)

        Return codifica
    End Function
    Public Function isValidUltima(ByVal cp As String) As String ' ok
        If cp.Trim = "" Then
            Return cp
        End If
        cp = cp.Trim
        Dim codifica As String = ""
        Dim s() As String = cp.Split(" ")
        Dim i As Integer = s.Length - 1

        codifica = s(i)

        Return codifica
    End Function
    Public Function isValidPrimeiraUltima(ByVal cp As String) As String ' ok
        If cp.Trim = "" Then
            Return cp
        End If
        Dim codifica As String = ""
        Dim s() As String = cp.Split(" ")
        Dim i As Integer = s.Length - 1

        codifica = s(0) & s(i)

        Return codifica
    End Function
    Public Function isValidMaiuscPP(ByVal cp As String) As String ' ok
        If cp.Trim = "" Then
            Return cp
        End If

        Dim codifica As String = ""
        Dim wn() As String = cp.Split(" ")

        For i As Integer = 0 To wn.Length - 1
            If i <> 0 Then
                codifica += " "
            End If
            If wn(i).Trim.ToUpper = "DE" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DA" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DO" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DES" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DAS" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DOS" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "A" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "E" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "I" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "O" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If
            If wn(i).Trim.ToUpper = "U" Then
                codifica += wn(i).Trim.ToLower
                Continue For
            End If

            Dim wuma As String = wn(i).Trim
            codifica += wuma.Substring(0, 1).ToUpper & wuma.Substring(1).ToLower
        Next

        Return codifica
    End Function
    Public Function isValidMaiuscTudo(ByVal cp As String) As String ' ok
        If cp.Trim = "" Then
            Return cp
        End If

        Return cp.ToUpper
    End Function
    Public Function isValidNomeAbreviado(ByVal cp As String) As String ' ok
        If cp.Trim = "" Then
            Return cp
        End If

        Dim codifica As String = ""
        Dim wn() As String = cp.Split(" ")

        Dim wprimeiro As String = ""
        Dim wmeio As String = ""
        Dim wultimo As String = ""

        Try
            wprimeiro = wn(0).Trim
            wprimeiro = wprimeiro.Substring(0, 1).ToUpper & wprimeiro.Substring(1).ToLower
        Catch ex As Exception
        End Try

        Try
            wultimo = wn(wn.Length - 1)
        Catch ex As Exception
        End Try

        If wn.Length = 0 Then
            wultimo = ""
        End If

        For i As Integer = 1 To wn.Length - 2
            If wn(i).Trim.ToUpper = "A" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "E" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "I" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "O" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "U" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DE" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DA" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DO" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DES" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DAS" Then
                Continue For
            End If
            If wn(i).Trim.ToUpper = "DOS" Then
                Continue For
            End If


            wmeio += wn(i).Substring(0, 1).ToUpper & "."
        Next

        codifica = wprimeiro & " " & wmeio & " " & wultimo

        Return codifica
    End Function
    Public Function isValidCodPost(ByVal cp As String) As Boolean 'ok
        If cp.Trim = "" Then
            Return True
        End If

        'cria o padrão regex
        Dim padraoRegex As String = "^\d{4}-\d{3}$"
        Dim verifica As New RegularExpressions.Regex(padraoRegex)
        'variavel boolean para tratar o status
        Dim valida As Boolean = False
        'verifica se o recurso foi fornecido
        If cp = "" Then
            'cep invalido
            valida = True
        Else
            'usa o método IsMatch Method para validar o regex
            valida = verifica.IsMatch(cp, 0)
        End If
        Return valida
    End Function
    Public Function isValidContNumerico(ByVal CntNum As String) As Boolean 'ok
        If CntNum.Trim = "" Then
            Return True
        End If

        If Not Regex.IsMatch(CntNum, "^[0-9]+$") Then
            Return False
        End If

        Return True
    End Function
End Class


