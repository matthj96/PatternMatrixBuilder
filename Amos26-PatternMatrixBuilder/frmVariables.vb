Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Public Class frmVariables

    Private Property hitOk As Boolean = False

    Public Function loadVariables(unfilteredVariables As ArrayList) As ArrayList

        'Load Unfiltered Variables Into ListBox + Select Them
        For Each variable As String In unfilteredVariables
            lstVariables.Items.Add(variable)
        Next

        Me.ShowDialog()

        Dim filteredVariables As New ArrayList

        If hitOk Then

            For i As Integer = 0 To lstVariables.SelectedItems.Count - 1
                filteredVariables.Add(lstVariables.SelectedItems(i))
            Next

        End If

        loadVariables = filteredVariables
    End Function

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        If criteriaCheck() Then
            hitOk = True
            Me.Close()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        hitOk = False
        Me.Close()
    End Sub

    Private Function criteriaCheck() As Boolean

        Dim filteredVariables As New ArrayList

        For i As Integer = 0 To lstVariables.SelectedItems.Count - 1
            filteredVariables.Add(lstVariables.SelectedItems(i))
        Next

        Dim factorVariables As New ArrayList
        Dim factors = New Dictionary(Of String, ArrayList)
        Dim factor As String

        For Each variable As String In filteredVariables

            Dim m As Match = Regex.Match(variable, "[^A-z]")

            If m.Success Then
                factor = variable.Substring(0, m.Index - 1)

                If factors.ContainsKey(factor) Then
                    factorVariables = factors.Item(factor)
                    factorVariables.Add(variable)
                    factors(factor) = factorVariables
                Else
                    factorVariables = New ArrayList
                    factorVariables.Add(variable)
                    factors.Add(factor, factorVariables)
                End If
            End If

        Next

        For Each factorDetail As KeyValuePair(Of String, ArrayList) In factors
            If factorDetail.Value.Count <= 1 Then
                MsgBox(factorDetail.Key & ": does not have more than one variable selected.")
                criteriaCheck = False
                Exit Function
            End If
        Next

        criteriaCheck = True
    End Function

    Private Sub btnToggleNone_Click(sender As Object, e As EventArgs) Handles btnToggleNone.Click
        For i As Integer = 0 To lstVariables.Items.Count - 1
            lstVariables.SetSelected(i, False)
        Next
    End Sub

    Private Sub btnToggleAll_Click(sender As Object, e As EventArgs) Handles btnToggleAll.Click
        For i As Integer = 0 To lstVariables.Items.Count - 1
            lstVariables.SetSelected(i, True)
        Next
    End Sub
End Class