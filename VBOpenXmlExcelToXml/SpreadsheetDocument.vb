Imports VBOpenXmlExcelToXml

Friend Class SpreadsheetDocument
    Friend Id As Object

    Friend Shared ReadOnly Property Open(filename As String, v As Boolean) As SpreadsheetDocument
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Friend Function WorkbookPart() As Object
        Throw New NotImplementedException()
    End Function
End Class
