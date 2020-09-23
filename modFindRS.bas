Attribute VB_Name = "modFindRS"
Option Explicit
'got this from PSC
'thanks
Public Function Search(parameter As String, rs As ADODB.Recordset, X As Field) As Boolean
    Dim foundFlag As Boolean
    Dim i As Integer

    With rs
        If .RecordCount > 0 Then
            .MoveFirst
                For i = 1 To .RecordCount
                    If X = parameter Then
                        foundFlag = True
                        i = .RecordCount
                    End If
                    If foundFlag = False Then
                        .MoveNext
                    End If
                Next i
                If foundFlag = True Then
                   'MsgBox ("Record has been location!")
                Else
                    MsgBox ("No Match in Database!")
                    foundFlag = False
                    .MoveFirst
                End If
        Else
            MsgBox ("There are no records To search!")
            foundFlag = False
        End If
    End With
    Search = foundFlag

End Function
