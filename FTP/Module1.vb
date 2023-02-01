Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Module Module1

    ' Объявление переменных для подключения к БД
    Public pathToDataBase As String = Application.StartupPath & "\db.mdb"
    Public connectToDataBase As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pathToDataBase)
    Public dataReader As OleDbDataReader

    ' Переменная для отчета
    Public report As String

    ' Переменные для авторизации
    Public userPermission As String = "Nothing"
    Public userPassword As String
    Public userIndex As String


    ' Процедура занесения записей в журнал
    Public Sub ZapGurnal()
        Try
            Dim Command As New OleDbCommand("INSERT INTO [Журнал] ([Запись]) values ('" & report & "')", connectToDataBase)
            connectToDataBase.Open()
            Command.ExecuteNonQuery()
            connectToDataBase.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            report = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    ' Процедура изменения данных в БД
    Public Sub changeNoteInDataBase(ByVal query As String)
        Try
            Dim Command As New OleDbCommand(query, connectToDataBase)
            connectToDataBase.Open()
            Command.ExecuteNonQuery()
            connectToDataBase.Close()
        Catch ex As Exception
            connectToDataBase.Close()
        End Try

    End Sub

    ' Процедура чтения данных из БД
    Public Sub readNoteFromDataBase(ByVal query As String)
        Dim command As New OleDbCommand(query, connectToDataBase)
        connectToDataBase.Open()
        dataReader = command.ExecuteReader

        While dataReader.Read = True
            userIndex = dataReader.GetValue(0)
            userPermission = dataReader.GetValue(1)
            userPassword = dataReader.GetValue(2)
        End While

        dataReader.Close()
        connectToDataBase.Close()
    End Sub

    ' Экспорт записей в Excel
    Public Sub exportDataToExcel(ByVal dataSource As Object)
        Dim myXL As Excel.Application, myWB As Excel.Workbook, myWS As Excel.Worksheet
        Dim i, y, z As Integer
        myXL = New Excel.Application
        myWB = myXL.Workbooks.Add
        myWS = myWB.Worksheets(1)
        z = 2
        myXL.Visible = True
        Try
            For i = 1 To dataSource.Items.Count

                For y = 1 To 6
                    myWS.Cells(1, y) = dataSource.Columns(y - 1).Text
                Next y

                For column = 0 To 5
                    myWS.Cells(z, column + 1) = dataSource.Items.Item(i - 1).SubItems.Item(column).Text
                Next column

                z = z + 1
            Next i

            For i = 1 To 6
                myWS.Columns(i).ColumnWidth = 20
            Next

            myXL = Nothing
            myWB = Nothing
            myWS = Nothing

            report = "Экспорт исходящих в Excel " & DateString & " " & TimeString : ZapGurnal()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            report = "Ошибка " & ex.Message & DateString & TimeString : ZapGurnal()
        End Try
    End Sub

    ' Экспорт записей в Word
    Public Sub exportDataToWord(ByVal dataSource As Object)
        Try
            Dim Дата As String = Format(Now, "d MMMM yyyy")
            Dim W = New Word.Application
            W.Visible = True
            W.Documents.Add()
            W.Selection.TypeText("Выписка из БД: " & Дата & Chr(13) & Chr(10))

            For i As Short = 0 To dataSource.Items.Count - 1
                W.Selection.TypeText(dataSource.Items(i).SubItems.Item(0).Text &
                                     " " & dataSource.Items(i).SubItems.Item(1).Text &
                                     " " & dataSource.Items(i).SubItems.Item(2).Text &
                                     " " & dataSource.Items(i).SubItems.Item(3).Text &
                                     " " & dataSource.Items(i).SubItems.Item(4).Text &
                                     " " & dataSource.Items(i).SubItems.Item(5).Text & Chr(13) & Chr(10))

            Next i
            W = Nothing
            report = "Экспорт исходящих в WORD " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            report = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

End Module
