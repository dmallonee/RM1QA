Imports Microsoft.VisualBasic

Public Class Utility

    


    Public Shared Function IsInt(ByVal strCheck As String, Optional ByVal CheckNegative As Boolean = False) As Boolean
        Dim re As Regex
        Dim m As Match

        If CheckNegative Then
            re = New Regex("^[\-]?[0-9]+$")
        Else
            re = New Regex("^[0-9]+$")
        End If

        If IsNothing(strCheck) Then
            Return False
            Exit Function
        End If

        m = re.Match(strCheck)

        Return m.Success

    End Function

    Public Shared Function FClean(ByVal strCheck As String) As String
        If Len(strCheck) > 0 Then
            Return strCheck
        Else
            Return ""
        End If
    End Function

    Public Shared Function StripPhone(ByVal strPhone As String) As String
        Dim strOutput As String
        strOutput = strPhone.Replace("(", "")
        strOutput = strOutput.Replace(")", "")
        strOutput = strOutput.Replace("-", "")
        strOutput = strOutput.Replace(" ", "")
        Return strOutput
    End Function

    Public Shared Function FormatPhone(ByVal strPhone As String) As String
        Dim strOutput As String
        If strPhone.Length > 9 Then
            strOutput = "(" & strPhone.Substring(0, 3) & ")" & strPhone.Substring(3, 3) & "-" & strPhone.Substring(6, 4)
        Else
            strOutput = strPhone
        End If
        
        Return strOutput
    End Function

    Public Shared Function IIFNotNull(ByVal DataItem As Object, Optional ByVal strInstead As String = "") As Object
        If Not (DataItem Is DBNull.Value) Then
            Return DataItem
        Else
            Return strInstead
        End If
    End Function


    Public Shared Function NotNull(ByVal DataItem As Object) As Boolean
        Return Not (DataItem Is DBNull.Value)
        
    End Function

    Public Shared Function IIFNotNull2(ByVal DataItem As Object, Optional ByVal Extra As String = "", Optional ByVal strInstead As String = "") As Object
        If Not (DataItem Is DBNull.Value) Then
            Return DataItem & Extra
        Else
            Return strInstead
        End If
    End Function

    Public Shared Function IsEmpty(ByVal DataItem As String, Optional ByVal strInstead As String = "") As Object
        If Len(DataItem) > 0 Then
            Return DataItem
        Else
            Return strInstead
        End If
    End Function

    ' Retrieves a web.config key of type string
    Public Shared Function XKey(ByVal key As String) As String
        Dim ASR As New AppSettingsReader
        Try
            Return CType(ASR.GetValue(key, GetType(System.String)), String)
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Shared Function Slash(ByVal sPath As String, Optional ByVal bLeadingSlash As Boolean = False) As String
        If sPath = "" Then
            Return ""
        Else
            sPath = Replace(sPath, "\", "/")

            If bLeadingSlash Then
                If Left(sPath, 1) = "/" Then
                    Return sPath
                Else
                    Return "/" & sPath
                End If
            Else
                Return sPath
            End If
        End If
    End Function

    Public Shared Function BSlash(ByVal sPath As String) As String
        sPath = Replace(sPath, "/", "\")
        If Left(sPath, 1) = "\" Then
            Return sPath
        Else
            Return "\" & sPath
        End If
    End Function

    Public Shared Function BSlashEnd(ByVal sPath As String) As String
        sPath = Replace(sPath, "/", "\")
        If Right(sPath, 1) = "\" Then
            Return sPath
        Else
            Return sPath & "\"
        End If
    End Function

    ' Checks for a specific item in a list of items
    Public Shared Function InList(ByVal strItem As String, ByVal strList As String) As Boolean
        Dim arrItems As String()
        Dim i As Integer
        Dim FoundIt As Boolean = False

        ' Redudant check to make sure the list is actually a list
        If InStr(strList, ",") > 0 Then

            arrItems = Split(strList, ",")
            For i = LBound(arrItems) To UBound(arrItems)
                If CStr(strItem) = CStr(arrItems(i)) Then
                    FoundIt = True
                End If
            Next

            ' This logic may not always work, but for now let's assume that the following call would return true:
            ' InList("hi","hi") just as if we called InList("hi","bye,hi,hello")... therefore...
        Else
            If CStr(strItem) = CStr(strList) Then FoundIt = True
        End If
        Return FoundIt
    End Function
End Class
