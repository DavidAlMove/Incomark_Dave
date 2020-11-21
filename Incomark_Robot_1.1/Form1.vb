Public Class Form1

    Dim p As Process
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = "0"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        'DATOS DEL INCOMARK:
        'PÁGINA WEB: https://www.sclpcj.com.mx:7071/SCLWeb/index.do
        'USUARIO: LT48012454
        'CONTRASEÑA: incomark13



        Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim dirArchivo As String = desktop + "\DatosAvales " + Date.Now.ToLongDateString + " - " + Date.Now.Hour.ToString() + "," + Date.Now.Minute.ToString() + ".xlsx"

        Dim dirArchivoStatic As String = dirArchivo

        'Threading.Thread.Sleep(4000)

        'Dim lll As String = "<td colspan=""2"">83263616</td>"
        'Dim unacaedenamas As String = ""
        'If lll.Contains("<td colspan=""2"">") Then
        '    Dim x As Integer = 0
        '    For Each t As Char In lll.ToCharArray()
        '        If IsNumeric(t) Then
        '            unacaedenamas = unacaedenamas + t
        '        End If
        '    Next
        'End If

        'unacaedenamas = unacaedenamas.Remove(0, 1)




        Dim dirDatos As String = desktop + "\IncomarkData.txt"

        Dim stringReader() As String = IO.File.ReadAllLines(dirDatos)

        Dim ii As Integer = 0

        ' Try
        Dim oExcel As Microsoft.Office.Interop.Excel.Application
        Dim oBook As Microsoft.Office.Interop.Excel.Workbook
        Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet


        oExcel = CreateObject("Excel.Application")
        oBook = oExcel.Workbooks.Add
        oSheet = oBook.ActiveSheet

        oSheet.Range("A1").Value = "Código Cliente"
        oSheet.Range("B1").Value = "Nombre"
        oSheet.Range("C1").Value = "Teléfono"
        oSheet.Range("D1").Value = "Nombre referencia 1"
        oSheet.Range("E1").Value = "Teléfono referencia 1"
        oSheet.Range("F1").Value = "Nombre referencia 2"
        oSheet.Range("G1").Value = "Teléfono referencia 2"
        oSheet.Range("H1").Value = "Nombre referencia 3"
        oSheet.Range("I1").Value = "Teléfono referencia 3"
        oSheet.Range("J1").Value = "Nombre referencia 4"
        oSheet.Range("K1").Value = "Teléfono referencia 4"

        Dim nombreHistorial As String
        Dim telefonoHistorial As String = "0"
        Dim nombreDatoPersonal1 As String = ""
        Dim telefonoDatoPersonal1 As String = ""
        Dim nombreDatoPersonal2 As String = ""
        Dim telefonoDatoPersonal2 As String = ""
        Dim nombreDatoPersonal3 As String = ""
        Dim telefonoDatoPersonal3 As String = ""
        Dim nombreDatoPersonal4 As String = ""
        Dim telefonoDatoPersonal4 As String = ""

        Dim arrayClave(100) As String
        Dim arrayNombreHistorial(100) As String
        Dim arrayTelefonoHistorial(100) As String
        Dim arrayNombreDatoPersonal1(100) As String
        Dim arrayTelefonoDatoPersonal1(100) As String
        Dim arrayNombreDatoPersonal2(100) As String
        Dim arrayTelefonoDatoPersonal2(100) As String
        Dim arrayNombreDatoPersonal3(100) As String
        Dim arrayTelefonoDatoPersonal3(100) As String
        Dim arrayNombreDatoPersonal4(100) As String
        Dim arrayTelefonoDatoPersonal4(100) As String

        If ii >= Integer.Parse(TextBox1.Text) Then
            GoTo terminar
        End If


        For Each clave As String In stringReader
            If ii >= Integer.Parse(TextBox1.Text) Then
                GoTo terminar
            End If
            stringReader(ii) = stringReader(ii).Replace(" ", "")
            Dim division As String() = stringReader(ii).Split("-")

            If division(0).Count < 2 Then
                division(0) = "0" + division(0)
            End If

            If division(1).Count < 2 Then
                division(1) = "0" + division(1)
            End If

            While (division(2).Count < 4)
                division(2) = "0" + division(2)
            End While

            While (division(3).Count < 4)
                division(3) = "0" + division(3)
            End While







            'stringReader(ii) = stringReader(ii).Replace("-", "")
            stringReader(ii) = division(0) + division(1) + division(2) + division(3)

            'Threading.Thread.Sleep(10000)

            Threading.Thread.Sleep(5000)
            SendKeys.Send(stringReader(ii))
            Threading.Thread.Sleep(2000)
            SendKeys.Send("{TAB}")
            Threading.Thread.Sleep(150)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(1000)
            SendKeys.Send("{ENTER}")
            If ii = 0 Then
                Threading.Thread.Sleep(14000)
            End If
            Threading.Thread.Sleep(9000)

            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            'REMOVE MENU
            SendKeys.Send("{F12}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f") 'CTRL + F
            Threading.Thread.Sleep(500)
            SendKeys.Send("ENTERADO")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 3}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^c") ' CTRL + C
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")

            Dim txtEnterado As String = Clipboard.GetText
            If txtEnterado.Trim = "¡Enterado!" Then
                Threading.Thread.Sleep(500)
                SendKeys.Send("{LEFT}")
                Threading.Thread.Sleep(500)
                SendKeys.Send("{ENTER 2}")
                Threading.Thread.Sleep(500)
                SendKeys.Send(" tabindex=""0""")
                Threading.Thread.Sleep(500)
                SendKeys.Send("{F12}")
                Threading.Thread.Sleep(500)
                SendKeys.Send("{TAB}")
                Threading.Thread.Sleep(500)
                SendKeys.Send("{ENTER}")
            End If
            'REMOVE MENU


            Threading.Thread.Sleep(500)
            SendKeys.Send("{F12}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^+C") 'CTRL + SHIFT + C
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB}")
            Threading.Thread.Sleep(500)

            SendKeys.Send("dropdown")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 3}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 3}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("class=""dropdown open"" tabindex=""0""")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")

            Threading.Thread.Sleep(500)
            SendKeys.Send("^f") 'CTRL + F
            Threading.Thread.Sleep(500)
            SendKeys.Send("{BACKSPACE 9}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("solicitud")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 3}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{LEFT}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{DEL 7}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("tabindex=""0"" onfocus")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")

            Threading.Thread.Sleep(500)
            SendKeys.Send("{F12}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB}")
            If ii = 0 Then
                Threading.Thread.Sleep(15000)
            End If
            Threading.Thread.Sleep(8000)

            SendKeys.Send("{TAB}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            If ii = 0 Then
                Threading.Thread.Sleep(10000)
            End If
            Threading.Thread.Sleep(4000)
            SendKeys.Send("{TAB}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            If ii = 0 Then
                Threading.Thread.Sleep(10000)
            End If
            Threading.Thread.Sleep(4000)
            SendKeys.Send("{TAB 9}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            If ii = 0 Then
                Threading.Thread.Sleep(5000)
            End If

            'OBTENER NOMBRE DE AVAL'
            SendKeys.Send("{F12}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f") 'CTRL + F
            Threading.Thread.Sleep(500)
            SendKeys.Send("Aval:")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 3}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 3}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^c") ' CTRL + C
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{F12}")
            Threading.Thread.Sleep(500)
            Dim copiadoNombreHistorial As String = Clipboard.GetText
            copiadoNombreHistorial = copiadoNombreHistorial.Trim
            'OBTENER NOMBRE DE AVAL'

            'OBTENER TELEFONO'
            SendKeys.Send("{F12}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f") 'CTRL + F
            Threading.Thread.Sleep(500)
            SendKeys.Send("Teléfono:")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 3}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 3}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^c") ' CTRL + C
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{F12}")
            Threading.Thread.Sleep(500)

            Dim copiadoTelefonoHistorial As String = Clipboard.GetText
            If copiadoTelefonoHistorial <> "td" Then
                If copiadoTelefonoHistorial.Contains("<td colspan=""2"">") Then
                    Dim x As Integer = 0
                    For Each t As Char In copiadoTelefonoHistorial.ToCharArray()
                        If IsNumeric(t) Then
                            telefonoHistorial = telefonoHistorial + t
                        Else
                            telefonoHistorial = "n/a"
                        End If

                    Next
                Else
                    copiadoTelefonoHistorial = "0N/A"
                End If

                If telefonoHistorial <> "" Then
                    telefonoHistorial = telefonoHistorial.Remove(0, 1)
                End If
            End If
            'OBTENER TELEFONO'


            'IR A "REFERENCIAS PERSONALES"
            SendKeys.Send("{TAB 9}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(3500)
            'IR A "REFERENCIAS PERSONALES"

            'SACAR NOMBRE 1 DE "DATOS PERSONALES"
            SendKeys.Send("{F12}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f") 'CTRL + F
            Threading.Thread.Sleep(500)
            SendKeys.Send("nombre")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(200)
            SendKeys.Send("{RIGHT}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(200)
            SendKeys.Send("{ENTER 2}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("^c")
            Threading.Thread.Sleep(250)

            Dim copiadoNombreDatoPersonal1 As String = Clipboard.GetText
            copiadoNombreDatoPersonal1 = copiadoNombreDatoPersonal1.Trim
            'SACAR NOMBRE 1 DE "DATOS PERSONALES"

            'SACAR EL TELEFONO 1 DE "DATOS PERSONALES"
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{BACKSPACE 6}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("Teléfono")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 4}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^c")
            Threading.Thread.Sleep(500)

            Dim copiadoTelefonoDatoPersonal1 As String = Clipboard.GetText
            copiadoTelefonoDatoPersonal1 = copiadoTelefonoDatoPersonal1.Trim

            SendKeys.Send("{ENTER}")
            'SACAR EL TELEFONO 1 DE "DATOS PERSONALES"

            'SACAR EL NOMBRE 2 DE "DATOS PERSONALES"
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f") 'CTRL + F
            Threading.Thread.Sleep(500)
            SendKeys.Send("{BACKSPACE 8}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("nombre")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 4}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(200)
            SendKeys.Send("{RIGHT}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(200)
            SendKeys.Send("{ENTER 2}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("^c")
            Threading.Thread.Sleep(250)

            Dim copiadoNombreDatoPersonal2 As String = Clipboard.GetText
            copiadoNombreDatoPersonal2 = copiadoNombreDatoPersonal2.Trim
            'SACAR NOMBRE 2 DE DATOS PERSONALES


            'SACAR TELEFONO 2 DE DATOS PERSONALES
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{BACKSPACE 6}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("Teléfono")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 4}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^c")
            Threading.Thread.Sleep(500)

            Dim copiadoTelefonoDatoPersonal2 As String = Clipboard.GetText
            copiadoTelefonoDatoPersonal2 = copiadoTelefonoDatoPersonal2.Trim

            SendKeys.Send("{ENTER}")
            'SACAR TELEFONO 2 DE DATOS PERSONALES


            'SACAR EL NOMBRE 3 DE "DATOS PERSONALES"
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f") 'CTRL + F
            Threading.Thread.Sleep(500)
            SendKeys.Send("{BACKSPACE 8}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("nombre")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 6}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(200)
            SendKeys.Send("{RIGHT}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(200)
            SendKeys.Send("{ENTER 2}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("^c")
            Threading.Thread.Sleep(250)

            Dim copiadoNombreDatoPersonal3 As String = Clipboard.GetText
            copiadoNombreDatoPersonal3 = copiadoNombreDatoPersonal3.Trim
            'SACAR NOMBRE 3 DE DATOS PERSONALES

            'SACAR TELEFONO 3 DE DATOS PERSONALES
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 3}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{BACKSPACE 6}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("Teléfono")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 4}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^c")
            Threading.Thread.Sleep(500)

            Dim copiadoTelefonoDatoPersonal3 As String = Clipboard.GetText
            copiadoTelefonoDatoPersonal3 = copiadoTelefonoDatoPersonal3.Trim

            SendKeys.Send("{ENTER}")
            'SACAR TELEFONO 3 DE DATOS PERSONALES

            ' - - - -- - 

            'SACAR EL NOMBRE 4 DE "DATOS PERSONALES"
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f") 'CTRL + F
            Threading.Thread.Sleep(500)
            SendKeys.Send("{BACKSPACE 8}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("nombre")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 8}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(200)
            SendKeys.Send("{RIGHT}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(200)
            SendKeys.Send("{ENTER 2}")
            Threading.Thread.Sleep(250)
            SendKeys.Send("^c")
            Threading.Thread.Sleep(250)

            Dim copiadoNombreDatoPersonal4 As String = Clipboard.GetText
            copiadoNombreDatoPersonal4 = copiadoNombreDatoPersonal4.Trim
            'SACAR NOMBRE 4 DE DATOS PERSONALES

            'SACAR TELEFONO 4 DE DATOS PERSONALES
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER 4}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^f")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{BACKSPACE 6}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("Teléfono")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 2}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{DOWN}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{TAB 4}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{ENTER}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^c")
            Threading.Thread.Sleep(500)

            Dim copiadoTelefonoDatoPersonal4 As String = Clipboard.GetText
            copiadoTelefonoDatoPersonal4 = copiadoTelefonoDatoPersonal4.Trim

            SendKeys.Send("{ENTER}")
            'SACAR TELEFONO 4 DE DATOS PERSONALES


            'SendKeys.Send("{F12}")
            'Threading.Thread.Sleep(500)
            'SendKeys.Send("{F5}")
            'Threading.Thread.Sleep(500)


            Dim XD As Integer = ii + 2
            oSheet.Range("A" + XD.ToString()).Value = clave
            oSheet.Range("B" + XD.ToString()).Value = copiadoNombreHistorial
            oSheet.Range("C" + XD.ToString()).Value = copiadoTelefonoHistorial
            oSheet.Range("D" + XD.ToString()).Value = copiadoNombreDatoPersonal1
            oSheet.Range("E" + XD.ToString()).Value = copiadoTelefonoDatoPersonal1
            oSheet.Range("F" + XD.ToString()).Value = copiadoNombreDatoPersonal2
            oSheet.Range("G" + XD.ToString()).Value = copiadoTelefonoDatoPersonal2
            oSheet.Range("H" + XD.ToString()).Value = copiadoNombreDatoPersonal3
            oSheet.Range("I" + XD.ToString()).Value = copiadoTelefonoDatoPersonal3
            oSheet.Range("J" + XD.ToString()).Value = copiadoNombreDatoPersonal4
            oSheet.Range("K" + XD.ToString()).Value = copiadoTelefonoDatoPersonal4



pasar:

            'Threading.Thread.Sleep(8000)
            'Threading.Thread.Sleep(8000)
            SendKeys.Send("{F12}")
            Threading.Thread.Sleep(500)
            SendKeys.Send("^1")
            Threading.Thread.Sleep(500)
            SendKeys.Send("{F5}")
            'Threading.Thread.Sleep(8000)
            ii = ii + 1


        Next

terminar:

        oBook.Close(True, dirArchivoStatic, False)
        oExcel.Quit()

        oSheet = Nothing
        oBook = Nothing
        oExcel = Nothing
        'Catch ex As Exception
        '   MsgBox("Hubo un error, guardando datos...")
        'End Try

        Dim ppp As Process() = Process.GetProcesses

        For Each p As Process In Process.GetProcesses
            If p.ProcessName = "EXCEL" Then p.Kill()
        Next


    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        'Kill("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")

    End Sub
End Class
