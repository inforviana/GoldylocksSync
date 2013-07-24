'GOLDYLOCKS SYNC
'Copyright Helder 'BurnSpirit' Correia @ 2012

Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports MySql.Data

Public Class Form1

    'nome e varsao da aplicacao
    Public Const versao As String = "GoldyLocks Sync 0.7a"

    'constantes para o EAN-13
    Private Const N As String = "N"
    Private Const A As String = "A"
    Private Const B As String = "B"
    Private Const C As String = "C"

    Dim W As String = "W"


    'funcao para calcular se desenha linha ou nao
    Private Function corlinha(ByVal digito As Integer, ByVal numero As Integer, ByVal posicao As Integer, ByVal numerolinha As Integer)
        Dim sequencia() As String, sequenciacor() As String, tipo As String

        Select Case digito
            Case 0
                sequencia = {12, A, A, A, A, A, A, C, C, C, C, C, C}
            Case 1
                sequencia = {12, A, A, B, A, B, B, C, C, C, C, C, C}
            Case 2
                sequencia = {12, A, A, B, B, A, B, C, C, C, C, C, C}
            Case 3
                sequencia = {12, A, A, B, B, B, A, C, C, C, C, C, C}
            Case 4
                sequencia = {12, A, B, A, A, B, B, C, C, C, C, C, C}
            Case 5
                sequencia = {12, A, B, B, A, A, B, C, C, C, C, C, C}
            Case 6
                sequencia = {12, A, B, B, B, A, A, C, C, C, C, C, C}
            Case 7
                sequencia = {12, A, B, A, B, A, B, C, C, C, C, C, C}
            Case 8
                sequencia = {12, A, B, A, B, B, A, C, C, C, C, C, C}
            Case 9
                sequencia = {12, A, B, B, A, B, A, C, C, C, C, C, C}
        End Select

        tipo = sequencia(posicao)

        Select Case numero
            Case 0
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, W, W, N, N, W, N}
                    Case B
                        sequenciacor = {7, W, N, W, W, N, N, N}
                    Case C
                        sequenciacor = {7, N, N, N, W, W, N, W}
                End Select
            Case 1
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, W, N, N, W, W, N}
                    Case B
                        sequenciacor = {7, W, N, N, W, W, N, N}
                    Case C
                        sequenciacor = {7, N, N, W, W, N, N, W}
                End Select
            Case 2
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, W, N, W, W, N, N}
                    Case B
                        sequenciacor = {7, W, W, N, N, W, N, N}
                    Case C
                        sequenciacor = {7, N, N, W, N, N, W, W}
                End Select
            Case 3
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, N, N, N, N, W, N}
                    Case B
                        sequenciacor = {7, W, N, W, W, W, W, N}
                    Case C
                        sequenciacor = {7, N, W, W, W, W, N, W}
                End Select
            Case 4
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, N, W, W, W, N, N}
                    Case B
                        sequenciacor = {7, W, W, N, N, N, W, N}
                    Case C
                        sequenciacor = {7, N, W, N, N, N, W, W}
                End Select
            Case 5
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, N, N, W, W, W, N}
                    Case B
                        sequenciacor = {7, W, N, N, N, W, W, N} ' {7, W, W, N, N, W, W, N}
                    Case C
                        sequenciacor = {7, N, W, W, N, N, N, W}
                End Select
            Case 6
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, N, W, N, N, N, N}
                    Case B
                        sequenciacor = {7, W, W, W, W, N, W, N}
                    Case C
                        sequenciacor = {7, N, W, N, W, W, W, W}
                End Select
            Case 7
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, N, N, N, W, N, N}
                    Case B
                        sequenciacor = {7, W, W, N, W, W, W, N}
                    Case C
                        sequenciacor = {7, N, W, W, W, N, W, W}
                End Select
            Case 8
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, N, N, W, N, N, N}
                    Case B
                        sequenciacor = {7, W, W, W, N, W, W, N}
                    Case C
                        sequenciacor = {7, N, W, W, N, W, W, W}
                End Select
            Case 9
                Select Case tipo
                    Case A
                        sequenciacor = {7, W, W, W, N, W, N, N}
                    Case B
                        sequenciacor = {7, W, W, N, W, N, N, N}
                    Case C
                        sequenciacor = {7, N, N, N, W, N, W, W}
                End Select

        End Select

        corlinha = sequenciacor(numerolinha)

    End Function

    Sub actualizar()
        Me.Enabled = False
        lstitems.Clear()

        Dim i As Integer = 0
        Dim querys As Integer = 0
        Dim c1, c2, c3, c4, c5, c6 As String
        Dim connstring As String
        Dim mysqlconn As New MySqlClient.MySqlConnection

        connstring = "server=" + txtservidor.Text + ";uid=" + txtusername.Text + ";pwd=" + txtpassword.Text + ";database=" + txtbd.Text + ";"
        mysqlconn.ConnectionString = "server=" + txtmysql.Text + ";user id=" + txtusermysql.Text + ";password=" + txtpasswordmysal.Text + ";database=" + txtbdmysql.Text


        Dim conn As New SqlConnection(connstring)

        conn.Open()
        mysqlconn.Open()

        Dim mysqlcmd As New MySqlClient.MySqlCommand
        Dim mysqlreader As MySqlClient.MySqlDataReader
        mysqlcmd.Connection = mysqlconn

        If chkartigos.Checked = True Then
            'actualizar artigos
            Dim cmd As New SqlCommand(txtquery.Text, conn)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            lstitems.Columns.Add("Referencia", 120)
            lstitems.Columns.Add("Descricao", 330)
            lstitems.Columns.Add("Stock Disp.", 120, HorizontalAlignment.Center)
            lstitems.Columns.Add("Preço", 120, HorizontalAlignment.Center)
            lstitems.Columns.Add("Estado", 120, HorizontalAlignment.Center)

            mysqlcmd.CommandText = "delete from artigos"
            mysqlcmd.ExecuteNonQuery()

            While reader.Read 'actualizar artigos
                lstitems.Items.Add(reader.GetString(0))
                lstitems.Items(i).SubItems.Add(reader.GetString(1))
                lstitems.Items(i).SubItems.Add(reader.GetDouble(2))
                lstitems.Items(i).SubItems.Add(reader.GetDouble(4))

                c1 = reader.GetString(1).Replace("'", "")
                c2 = reader.GetString(0).Replace("'", "")
                c3 = reader.GetInt64(3)
                c4 = reader.GetDouble(4)
                c5 = reader.GetInt32(5)

                If (c5 = 0) Then
                    mysqlcmd.CommandText = "insert into artigos (nome,cod_barras,id_familia,preco_custo) values ('" + c1 + "','" + c2 + "'," + c3 + "," + c4.Replace(",", ".") + ")"
                End If

                If (c5 = 1) Then
                    mysqlcmd.CommandText = "update artigos set preco = '" + c4.Replace(",", ".") + "' where cod_barras = '" + c2 + "'"
                End If

                Try
                    mysqlcmd.ExecuteNonQuery()
                    lstitems.Items(i).SubItems.Add("Actualizado")
                Catch ex As Exception
                    lstitems.Items(i).SubItems.Add("Erro ao actualizar!")
                End Try
                i = i + 1
                querys = querys + 1
            End While
            reader.Close()
        End If


        If chkfamilias.Checked = True Then
            'actualizar familias
            Dim cmdf As New SqlCommand("select FamilyID, Description, ParentDescription from family", conn) 'query para ler as familias no sql server
            Dim readerf As SqlDataReader = cmdf.ExecuteReader() 'executar a query

            mysqlcmd.CommandText = "delete from familias" 'query para apagar as familias
            mysqlcmd.ExecuteNonQuery() 'executa a query

            While readerf.Read() 'ciclo
                c1 = readerf.GetInt64(0)
                c2 = readerf.GetString(1)
                c3 = readerf.GetString(2)

                'guardar as familias no mysql
                mysqlcmd.CommandText = "insert into familias (id_familia, descricao, familia_pai) values (" + c1 + ",'" + c2 + "','" + c3 + "')"
                mysqlcmd.ExecuteNonQuery()
                querys = querys + 1
            End While
            readerf.Close()
        End If



        If chkclientes.Checked = True Then
            'actualizar clientes 
            Dim cmdc As New SqlCommand("select CustomerID, OrganizationName, FederalTaxID, DirectDiscount from customer", conn)
            Dim readerc As SqlDataReader = cmdc.ExecuteReader

            mysqlcmd.CommandText = "delete from clientes"
            mysqlcmd.ExecuteNonQuery()

            While readerc.Read
                c1 = readerc.GetInt64(0)
                c2 = readerc.GetString(1).Replace("'", "")
                c3 = readerc.GetString(2).Replace("'", "")
                c4 = readerc.GetDouble(3)


                mysqlcmd.CommandText = "insert into clientes (cod_sage, nome, nif,desconto_global) values (" + c1 + ",'" + c2 + "','" + c3 + "'," + c4 + ")"
                mysqlcmd.ExecuteNonQuery()
                querys = querys + 1
            End While

            readerc.Close()
        End If


        If chkcc.Checked Then
            Dim cmdcc As New SqlCommand("select customerledgeraccount.PartyID, Sum(CustomerLedgerAccount.TotalAmount) as 'Total' from CustomerLedgerAccount group by CustomerLedgerAccount.PartyID;")
            Dim readercc As SqlDataReader = cmdcc.ExecuteReader

            mysqlcmd.CommandText = "delete from clientes_cc"
            mysqlcmd.ExecuteNonQuery()

            While readercc.Read
                c1 = readercc.GetInt64(0)
                c2 = readercc.GetDouble(1)

                mysqlcmd.CommandText = "insert into clientes_cc (id_cliente, valor_cc) VALUES (" + c1 + "," + c2 + ")"
                mysqlcmd.ExecuteNonQuery()
            End While

            readercc.Close()
        End If


        If chkdescontos.Checked = True Then
            'actualizar descontos
            Dim cmdd As New SqlCommand("select discountid, itemid, familyid, customerid, unitprice, taxincludedprice from discounteligibleitem", conn)
            Dim readerd As SqlDataReader = cmdd.ExecuteReader

            mysqlcmd.CommandText = "delete from descontos_artigos" 'eliminar descontos nos artigo
            mysqlcmd.ExecuteNonQuery()

            mysqlcmd.CommandText = "delete from descontos" 'eliminar descontos especificos nos artigos
            mysqlcmd.ExecuteNonQuery()

            While readerd.Read 'ciclo descontos artigos
                c1 = readerd.GetInt32(0)
                c2 = readerd.GetString(1).Replace("'", "")
                c3 = readerd.GetInt64(2)
                c4 = readerd.GetInt64(3)
                c5 = readerd.GetDouble(4)
                c6 = readerd.GetDouble(5)

                c5 = c5.Replace(",", ".") 'FIX casas decimais no preco

                mysqlcmd.CommandText = "insert into descontos_artigos (id_descontos_artigos, id_artigo, id_familia, id_cliente, preco) values (" + c1 + ",'" + c2 + "'," + c3 + "," + c4 + ",'" + c5 + "')"
                mysqlcmd.ExecuteNonQuery()
            End While

            readerd.Close()

            Dim cmddd As New SqlCommand("select discountid, discount1 from discountdefinition", conn)
            Dim readerdd As SqlDataReader = cmddd.ExecuteReader

            While readerdd.Read 'ciclo descontos
                c1 = readerdd.GetInt32(0)
                c2 = readerdd.GetDouble(1)

                mysqlcmd.CommandText = "insert into descontos (id_desconto, desconto) values (" + c1 + "," + c2 + ")"
                mysqlcmd.ExecuteNonQuery()
            End While
        End If

        mysqlconn.Close()
        conn.Close()
        Me.Enabled = True
    End Sub

    'ler configuracao
    Sub ler_config()
        'titulo da janela
        Me.Text = versao

        Dim linha() As String 'linha a ser lida do ficheiro
        Dim conf, val As String 'valores a guardar da linha

        If (File.Exists("config.ini")) Then 'verifica se existe o ficheiro
            Dim ficheiro_config As StreamReader = New StreamReader("config.ini") 'abre o ficheiro para leitura
            While ficheiro_config.Peek <> -1 'le ate ao fim do ficheiro
                linha = ficheiro_config.ReadLine().Split("=")
                conf = linha(0)
                val = linha(1)
                If (conf = "SQLSERVER_SERVER") Then
                    txtservidor.Text = val
                ElseIf (conf = "SQLSERVER_USER") Then
                    txtusername.Text = val
                ElseIf (conf = "SQLSERVER_PASSWORD") Then
                    txtpassword.Text = val
                ElseIf (conf = "SQLSERVER_BD") Then
                    txtbd.Text = val
                ElseIf (conf = "MYSQLSERVER_SERVER") Then
                    txtmysql.Text = val
                ElseIf (conf = "MYSQLSERVER_USER") Then
                    txtusermysql.Text = val
                ElseIf (conf = "MYSQLSERVER_PASSWORD") Then
                    txtpasswordmysal.Text = val
                ElseIf (conf = "MYSQLSERVER_BD") Then
                    txtbdmysql.Text = val
                ElseIf (conf = "ETIQUETAS_PRINTER") Then
                    cmbimpressoras.Text = val
                ElseIf (conf = "ETIQUETAS_TAMANHO") Then
                    cmbtamanho.Text = val
                ElseIf (conf = "UPDATE_TIME") Then
                    cmbupdatetime.Text = val
                    If (val.Length > 0) Then
                        chkupdate.Checked = True
                    End If
                ElseIf (conf = "UPDATE") Then
                    If val = "1" Then
                        chkupdate.Checked = True
                    Else
                        chkupdate.Checked = False
                    End If
                End If

            End While
            ficheiro_config.Close()
        Else 'se nao existir o ficheiro cria um novo
            Dim ficheiro_config As StreamWriter = New StreamWriter("config.ini")
            ficheiro_config.WriteLine("SQLSERVER_SERVER=")
            ficheiro_config.WriteLine("SQLSERVER_USER=")
            ficheiro_config.WriteLine("SQLSERVER_PASSWORD=")
            ficheiro_config.WriteLine("SQLSERVER_BD=")
            ficheiro_config.WriteLine("MYSQLSERVER_SERVER=")
            ficheiro_config.WriteLine("MYSQLSERVER_USER=")
            ficheiro_config.WriteLine("MYSQLSERVER_PASSWORD=")
            ficheiro_config.WriteLine("MYSQLSERVER_BD=")
            ficheiro_config.WriteLine("UPDATE_TIME=")
            ficheiro_config.WriteLine("UPDATE=")
            ficheiro_config.Close()
        End If
        agendar()
    End Sub

    'guardar configuracao
    Sub escrever_config()
        Dim ficheiro_config As StreamWriter = New StreamWriter("config.ini", False)
        ficheiro_config.WriteLine("SQLSERVER_SERVER="+txtservidor.Text)
        ficheiro_config.WriteLine("SQLSERVER_USER=" + txtusername.Text)
        ficheiro_config.WriteLine("SQLSERVER_PASSWORD=" + txtpassword.Text)
        ficheiro_config.WriteLine("SQLSERVER_BD=" + txtbd.Text)
        ficheiro_config.WriteLine("MYSQLSERVER_SERVER=" + txtmysql.Text)
        ficheiro_config.WriteLine("MYSQLSERVER_USER=" + txtusermysql.Text)
        ficheiro_config.WriteLine("MYSQLSERVER_PASSWORD=" + txtpasswordmysal.Text)
        ficheiro_config.WriteLine("MYSQLSERVER_BD=" + txtbdmysql.Text)
        ficheiro_config.WriteLine("UPDATE_TIME=" + cmbupdatetime.Text)
        ficheiro_config.WriteLine("ETIQUETAS_PRINTER=" + cmbimpressoras.Text)
        ficheiro_config.WriteLine("ETIQUETAS_TAMANHO=" + cmbtamanho.Text)
        If chkupdate.Checked = True Then
            ficheiro_config.WriteLine("UPDATE=1")
        Else
            ficheiro_config.WriteLine("UPDATE=0")
        End If
        ficheiro_config.Close()
    End Sub

    'agendar sincronizaca
    Sub agendar()
        Dim tempo As Integer
        If (chkupdate.Checked = True And cmbupdatetime.Text.Length > 0) Then
            tempo = CInt(cmbupdatetime.Text)
            autoupdate.Interval = (tempo * 60) * 1000
            autoupdate.Enabled = True
            lblupdate.Text = "Agendamento Activo..." + autoupdate.Interval.ToString
        Else
            autoupdate.Enabled = False
            lblupdate.Text = "Agendamento Desactivado"
        End If
    End Sub
    Private Sub obter_impressoras()
        'obter lista de impressoras do sistema
        Dim impressoras As String

        For Each impressoras In Printing.PrinterSettings.InstalledPrinters
            cmbimpressoras.Items.Add(impressoras)
        Next
    End Sub
    Private Sub detalhes_impressora()
        'obter configuracoes da impressora do tamanho do papel
        Dim tamanhos As Integer
        tamanhos = printer.PrinterSettings.PaperSizes.Count

        'limpar os tamanhos da combo
        cmbtamanho.Items.Clear()

        'mostra os tamanhos disponiveis
        For i As Integer = 0 To tamanhos - 1
            cmbtamanho.Items.Add(printer.PrinterSettings.PaperSizes.Item(i).PaperName.ToString)
        Next
    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'obter as impressoras do sistema
        obter_impressoras()

        'carregar detalhes da impressora
        detalhes_impressora()

        'carrega o ficheiro de configuracao
        ler_config()
    End Sub

    Private Sub SairToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SairToolStripMenuItem1.Click
        End
    End Sub

    Private Sub SobreToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SobreToolStripMenuItem1.Click
        MsgBox(versao + vbNewLine + "Inforviana 2013")
    End Sub

    Private Sub cmdligar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        escrever_config()
        agendar()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        escrever_config()
        actualizar()
    End Sub

    Private Sub txtquery_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbupdatetime_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub autoupdate_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles autoupdate.Tick
        lblupdate.Text = "A sincronizar..."
        escrever_config()
        actualizar()
        lblupdate.Text = "Agendamento Activo..." + autoupdate.Interval.ToString
    End Sub

    Private Sub chkupdate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Form1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Leave
        escrever_config()
    End Sub

    Private Sub printer_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles printer.PrintPage
        Dim rect As New Rectangle(1, 5, 170, 250)
        e.Graphics.DrawImage(PEan.Image, rect)
    End Sub

    Private Sub etiquetas(ByVal codigo As String, ByVal qtd As Integer, ByVal descricao1 As String, ByVal descricao2 As String, ByVal local As String)
        PEan.Image = Nothing 'limpa a picture box

        Dim inicial As Integer
        Dim resto As String
        Dim grupo, x, x1, coluna, numerodegrupo, pposicao, nnumero, currentX, currentY As Integer
        Dim alturamax = 80 'altura do codigo de barras

        If (codigo.Length = 12) Then 'calcular digito de verificacao
            Dim soma_inpar As Integer = 0
            Dim soma_par, soma_total As Integer
            Dim inpar_mult As Integer
            Dim divisao As Integer
            Dim digito As Integer
            soma_par = 0


            soma_inpar = CInt(Mid(codigo, 12, 1)) + CInt(Mid(codigo, 10, 1)) + CInt(Mid(codigo, 8, 1)) + CInt(Mid(codigo, 6, 1)) + CInt(Mid(codigo, 4, 1)) + CInt(Mid(codigo, 2, 1))
            soma_par = CInt(Mid(codigo, 11, 1)) + CInt(Mid(codigo, 9, 1)) + CInt(Mid(codigo, 7, 1)) + CInt(Mid(codigo, 5, 1)) + CInt(Mid(codigo, 3, 1)) + CInt(Mid(codigo, 1, 1))
            inpar_mult = soma_inpar * 3
            soma_total = inpar_mult + soma_par
            If (soma_total Mod 10 = 0) Then
                codigo = codigo + "0"
            Else
                divisao = soma_total / 10
                digito = ((divisao + 1) * 10) - soma_total
                If (digito > 9) Then
                    digito = digito - 10 'FIX Norberto
                End If
                codigo = codigo + digito.ToString
            End If
        Else
            'MessageBox.Show("Erro CB:" + codigo)
        End If

        If (IsNumeric(codigo) And codigo.Length = 13) Then 'verifica se so tem numeros
            inicial = Mid(codigo, 1, 1) 'numero inicial
            resto = Mid(codigo, 2, 12) 'resto dos numeros

            Dim bmp As New Bitmap(500, 500)
            Dim g As Graphics = Graphics.FromImage(bmp)
            Dim p As Pen = New Pen(Color.Black, 3)

            g.DrawLine(p, 56, 10, 56, alturamax) 'linhas iniciais 
            g.DrawLine(p, 62, 10, 62, alturamax)

            g.DrawString(inicial.ToString, lbletiqueta.Font, Brushes.Black, 20, alturamax - 20) 'primeiro algarismo
            currentX = 62
            For grupo = 1 To 2 'divide o codigo em conjuntos de seis algarismos
                Select Case grupo
                    Case 1
                        x = 80
                        x1 = 80
                    Case 2
                        x = 400
                        x1 = 400
                End Select
                For numerodegrupo = 1 To 6 'algarismos
                    pposicao = IIf(grupo = 1, numerodegrupo, numerodegrupo + 6)
                    nnumero = IIf(grupo = 1, Mid(resto, numerodegrupo, 1), Mid(resto, numerodegrupo + 6, 1))
                    For coluna = 1 To 7 'barras de cada algarismo
                        currentX = currentX + 3
                        If coluna = 1 Then
                            currentY = alturamax - 20 'posicao do texto
                            g.DrawString(nnumero, lbletiqueta.Font, Brushes.Black, currentX, currentY)
                        End If
                        If (corlinha(inicial, nnumero, pposicao, coluna) = "N") Then
                            g.DrawLine(p, currentX, 10, currentX, alturamax - 20)
                        End If
                    Next
                Next
                If (grupo = 1) Then
                    currentX = currentX + 3

                End If
                currentX = currentX + 3
                g.DrawLine(p, currentX, 10, currentX, alturamax)
                currentX = currentX + 6
                g.DrawLine(p, currentX, 10, currentX, alturamax)
                currentX = currentX + 3

                If (chkreferencia.Checked = True) Then
                    g.DrawString("Ref. " + descricao1, lbletiqueta.Font, Brushes.Black, 20, alturamax + 10) 'descricao 1
                End If

                Dim maxd1 As Integer

                If (chkdescricao.Checked = True) Then
                    If (descricao2.Length > 27) Then '1 linha da descricao 2
                        maxd1 = 27
                    Else
                        maxd1 = descricao2.Length
                    End If

                    g.DrawString(descricao2.Substring(0, maxd1), lblfont1.Font, Brushes.Black, 20, alturamax + 40) 'descricao 2

                    If (descricao2.Length > 27) Then '2 linha da descricao 2
                        Try
                            g.DrawString(descricao2.Substring(27, descricao2.Length - 27), lblfont1.Font, Brushes.Black, 20, alturamax + 60)
                        Catch ex As Exception
                        End Try
                        'descricao 2 continuada
                    End If
                End If
                If (chklocalizacao.Checked = True) Then
                    g.DrawString("Local :  " + local, lbletiqueta.Font, Brushes.Black, 20, alturamax + 80) 'posicao na prateleira
                End If

                PEan.Image = bmp
            Next


            printer.PrinterSettings.PrinterName = cmbimpressoras.Text  'impressora para as etiquetas

            Try
                'If cmbtamanho.SelectedIndex <> -1 Then
                printer.DefaultPageSettings.PaperSize = printer.PrinterSettings.PaperSizes.Item(cmbtamanho.SelectedIndex)
                For i = 1 To qtd
                    printer.Print() 'imprime a etiqueta para  a impressora pre definida
                Next
                'End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Else
            Me.Text = "2"
            'MessageBox.Show("Erro CB 2:" + codigo)
        End If
    End Sub

    Private Sub tmr_etiquetas_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmr_etiquetas.Tick
        If (chketiquetas.Checked = True) Then
            Dim mysqlconn As New MySqlClient.MySqlConnection
            mysqlconn.ConnectionString = "server=" + txtmysql.Text + ";user id=" + txtusermysql.Text + ";password=" + txtpasswordmysal.Text + ";database=" + txtbdmysql.Text
            mysqlconn.Open()
            Dim mysqlcmd As New MySqlClient.MySqlCommand
            Dim mysqlreader As MySqlClient.MySqlDataReader
            mysqlcmd.Connection = mysqlconn
            mysqlcmd.CommandText = "select spooler.ean,artigos_cb.id_artigo,artigos.nome, artigos_posicao.posicao, spooler.qtd from spooler left join artigos_cb on artigos_cb.codigo=spooler.ean left join artigos on artigos.cod_barras=artigos_cb.id_artigo left join artigos_posicao on artigos_posicao.cod_barras = artigos_cb.id_artigo"
            mysqlreader = mysqlcmd.ExecuteReader()
            If mysqlreader.HasRows Then
                While mysqlreader.Read
                    Try
                        etiquetas(mysqlreader.GetValue(0), mysqlreader.GetString(4), mysqlreader.GetString(1), mysqlreader.GetString(2), mysqlreader.GetString(3))
                    Catch
                        Try
                            etiquetas(mysqlreader.GetValue(0), mysqlreader.GetString(4), mysqlreader.GetString(1), mysqlreader.GetString(2), "")
                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)

                        End Try
                    End Try
                End While
            End If
            mysqlreader.Close()
            mysqlcmd.CommandText = "delete from spooler"
            mysqlcmd.ExecuteNonQuery()
            mysqlconn.Close()
        End If
    End Sub

    Private Sub cmdetiq_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdetiq.Click
        escrever_config()
    End Sub

    Private Sub txtquery_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtquery.TextChanged

    End Sub

    Private Sub AjudaToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AjudaToolStripMenuItem1.Click

    End Sub

    Private Sub cmbimpressoras_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbimpressoras.SelectedIndexChanged
        printer.PrinterSettings.PrinterName = cmbimpressoras.Text 'impressora para as etiquetas
        detalhes_impressora()
    End Sub
End Class