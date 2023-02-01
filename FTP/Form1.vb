Imports System.ComponentModel
Imports System.IO

Public Class Form1

    'Переменные для распределения данных с логов
    Public logData, logTime, logProtocol, logName, logNote, logStatus, logIp, searchTxt As String

    'Переменная для запроса в БД
    Public query As String

    'Переменные для поиска записей в ListView
    Public criterion As String
    Public searchColumn As Integer

    'Переменная для определения состояния входа пользователя в аккаунт
    Public isLogin As Boolean = False


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        startProgram()
    End Sub


    'Процедура инициализации программы: шаг 1
    Private Sub startProgram()
        ListView1.Items.Clear()
        ToolStrip2.Visible = False
        ListView1.Visible = False
        report = "Программа успешно запущена. " & "Пользователь: " & userPermission & " " & DateString & " " & TimeString : ZapGurnal()
        ToolStripLabel5.Text = "Авторизированы права: " & userPermission
    End Sub

    'Кнопка для входа пользователя в аккаунт
    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If isLogin = False Then
            Call login()
        Else
            ToolStripButton3.Text = "Войти"
            userPermission = "Nothing"
            ToolStripLabel5.Text = "Авторизированы права: " & userPermission
            report = "Выход пользователя " & userPermission & " " & DateString & " " & TimeString : ZapGurnal()
            rebootProgram()
            isLogin = False
        End If
    End Sub

    'Процедура очистки ListView и параметров поиска записи
    Private Sub rebootProgram()
        ToolStripComboBox1.Text = ""
        TextBox1.Text = ""
        deleteReadLogsFromDataBase()
        readDataFromLogFiles()
    End Sub

    'Очистка временной таблицы "Логи" в БД
    Private Sub deleteReadLogsFromDataBase()
        query = "DELETE FROM Логи"
        changeNoteInDataBase(query)

        ListView1.Columns.Clear()
        ListView1.Items.Clear()
    End Sub


    'Процедура аутентификации пользователя: шаг 2
    Private Sub login()
        Dim verificationWindow As String ' Окно авторизации

        Try
            verificationWindow = InputBox("Введите пароль для авторизации пользователя", "Запрос авторизации")

            If verificationWindow = "" Then
                Exit Sub
            Else

                ' Запрос на получение пользователя по введенному паролю
                query = "Select Пользователи.[Код], Пользователи.[Права], Пользователи.[Пароль] FROM Пользователи WHERE (((Пользователи.[Пароль]) ='" &
                    verificationWindow & "'));"

                ' Вызов процедуры на выполнение запроса в модуле
                readNoteFromDataBase(query)

                ' Обработка результата запроса:
                If userIndex = "" Then
                    MessageBox.Show("Авторизация не удалась", "Неудача", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Else

                    ' В случае входа:
                    ToolStrip2.Visible = True
                    ListView1.Visible = True
                    ToolStripLabel5.Text = "Авторизированы права: " & userPermission
                    Call readDataFromLogFiles() ' Вызов процедуры чтения лог-файлов
                    isLogin = True
                End If

                report = "Авторизация в программе " & DateString & " " & TimeString : ZapGurnal()
                report = "Пользователь: " & userPermission & " " & DateString & " " & TimeString : ZapGurnal()

            End If

            ToolStripButton3.Text = "Выйти"

        Catch ex As Exception
            connectToDataBase.Close()
            MsgBox(ex.Message)
            report = "Ошибка авторизации: " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub



    Private Sub ЭкспортWORDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортWordToolStripMenuItem.Click
        exportDataToWord(ListView1)
    End Sub

    Private Sub ЭкспортExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортExcelToolStripMenuItem.Click
        exportDataToExcel(ListView1)
    End Sub

    Private Sub ЭкспортAccessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортAccessToolStripMenuItem.Click
        loadDataToDataBase()
    End Sub

    ' Загрузка записей ListView в Access
    Private Sub loadDataToDataBase()
        Try

            For i = 0 To ListView1.Items.Count - 1
                logData = ListView1.Items(i).SubItems(0).Text
                logTime = ListView1.Items(i).SubItems(1).Text
                logProtocol = ListView1.Items(i).SubItems(2).Text
                logName = ListView1.Items(i).SubItems(3).Text
                logStatus = ListView1.Items(i).SubItems(4).Text
                logIp = ListView1.Items(i).SubItems(5).Text

                query = "Insert Into [Логи] ([Дата], [Время], [Протокол], [Логин], [Статус], [IP]) values ('" & logData & "', '" &
                logTime & "','" & logProtocol & "','" & logName & "','" & logStatus & "','" & logIp & "')"
                changeNoteInDataBase(query)
            Next

        Catch ex As Exception

        End Try
    End Sub

    ' Обновление даты/времени
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ToolStripLabel2.Text = Now.ToLongDateString
        ToolStripLabel3.Text = Now.ToLongTimeString
    End Sub

    ' Поиск записи 
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        searchNote()
    End Sub

    ' Чтение данных с лог-файлов
    Private Sub readDataFromLogFiles()

        ' Верификация прав доступа
        If userPermission = "Admin" Or userPermission = "Developer" Or userPermission = "Tester" Then

            ' Очистка и создание столбцов ListView
            With ListView1
                .Columns.Clear()
                .Columns.Add("Дата")
                .Columns.Add("Время")
                .Columns.Add("Протокол")
                .Columns.Add("Логин")
                .Columns.Add("Статус")
                .Columns.Add("IP")
                .Items.Clear()
            End With

            ' Переменные для получение списка файлов в каталоге и их кол-ва
            Dim files = New DirectoryInfo(Application.StartupPath & "\Logs").GetFiles.OrderBy(Function(x) Val(Path.GetFileNameWithoutExtension(x.Name)))
            Dim countLogFiles As Integer


            ' Перебор всех файлов в каталоге
            For Each file In files

                ' Получение имени и пути файлов
                Dim filePath As String = file.FullName
                Debug.Print($"Filepath: {filePath}")

                ' Объявление "читателя"
                Dim inputStream As New StreamReader(filePath)
                Dim newStream() As String = Nothing

                ' Чтение всех строк с разделителем
                Do While inputStream.Peek <> -1
                    newStream = inputStream.ReadLine().Split(" ")
                    logData = newStream(0)
                    logTime = newStream(1)
                    logProtocol = newStream(2)
                    logName = newStream(3)
                    logNote = newStream(4)
                    logStatus = newStream(5)
                    logIp = newStream(6)

                    With ListView1
                        .Items.Add(logData)
                        .Items.Item(.Items.Count - 1).SubItems.Add(logTime)
                        .Items.Item(.Items.Count - 1).SubItems.Add(logProtocol)
                        .Items.Item(.Items.Count - 1).SubItems.Add(logName)
                    End With

                    If logStatus = "in," Then

                        logStatus = "logged in"

                        With ListView1
                            .Items.Item(.Items.Count - 1).SubItems.Add(logStatus)
                            .Items.Item(.Items.Count - 1).SubItems.Add(logIp)
                        End With

                        Debug.Print($"{logData}|{logTime}|{logProtocol}|{logName}|{logNote}|{logStatus}|{logIp}")

                    Else

                        If logStatus = "out," Then

                            logStatus = "logged out"
                            logIp &= $" {newStream(7)} {newStream(8)} {newStream(9)} {newStream(9)}"

                            With ListView1
                                .Items.Item(.Items.Count - 1).SubItems.Add(logStatus)
                                .Items.Item(.Items.Count - 1).SubItems.Add(logIp)
                            End With

                            Debug.Print($"{logData}|{logTime}|{logProtocol}|{logName}|{logNote}|{logStatus}|{logIp}")

                        End If
                    End If
                Loop
                inputStream.Close()
            Next

            countLogFiles = New DirectoryInfo(Application.StartupPath & "\Logs\").GetFiles("*.txt", SearchOption.TopDirectoryOnly).Count.ToString
            ToolStripLabel1.Text = $"Количество записей: {ListView1.Items.Count}; Количество файлов: {countLogFiles}"
            ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
        End If
    End Sub

    Private Sub searchNote()

        If ToolStripComboBox1.SelectedIndex <> -1 Then
            If TextBox1.Text.Length <> 0 Then
                criterion = TextBox1.Text
                searchColumn = ToolStripComboBox1.SelectedIndex
                Dim i As Integer = 0
                Try
                    While i <> ListView1.Items.Count
                        If ListView1.Items(i).SubItems(searchColumn).Text.Contains(criterion) = False Then
                            ListView1.Items(i).Remove()
                            i -= 1
                        Else
                            Debug.Print($"Search note: {i}. {ListView1.Items(i).SubItems(searchColumn).Text}")
                        End If
                        i += 1
                    End While
                Catch ex As Exception

                End Try
            End If
        End If

        ToolStripLabel1.Text = "Количество записей: " & ListView1.Items.Count

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        rebootProgram()
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        deleteReadLogsFromDataBase()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then Call searchNote()
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F5 Then
            rebootProgram()
        End If
        If e.KeyCode = Keys.Escape Then
            Application.Exit()
        End If
    End Sub
End Class
