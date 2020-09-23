Imports System
Imports Microsoft.VisualBasic
Imports Amos
Imports AmosEngineLib
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports MiscAmosTypes
Imports MiscAmosTypes.cDatabaseFormat
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

<System.ComponentModel.Composition.Export(GetType(Amos.IPlugin))>
Public Class CustomCode
    Implements IPlugin

    Private Property factors As Dictionary(Of String, ArrayList)
    Private Property filteredVariables As ArrayList
    Private Property smartNaming As Boolean = True
    Private Property pageWidth As Double = pd.PageWidth
    Private Property pageHeight As Double = pd.PageHeight
    Private Property verticalSeparation As Double
    Private Property fontSize As Double
    Private Property estimatesChecked As Boolean

    Public Function Name1() As String Implements IPlugin.Name
        Return "CFA Auto Modeler"
    End Function

    Public Function Description() As String Implements IPlugin.Description
        Return "Draws an Amos model from an SPSS data file."
    End Function

    Public Function MainSub() As Integer Implements IPlugin.MainSub
        While DataFileIsEmpty() = True
            If MsgBox("Please specify a data file to continue.", MsgBoxStyle.OkCancel) = 2 Then
                Return 0
            End If
            pd.FileDataFiles()
        End While

        ParseData()
    End Function

    Private Function DataFileIsEmpty() As Boolean
        Dim dbFormat As cDatabaseFormat
        Dim fileName As String = ""
        Dim tableName As String = ""
        Dim groupingVariable As String = ""
        Dim groupingValue As Object = vbNull
        pd.GetDataFile(1, dbFormat, fileName, tableName, groupingVariable, groupingValue)

        If IsNullOrWhiteSpace(fileName) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsNullOrWhiteSpace(inputString As String) As Boolean
        If String.IsNullOrEmpty(inputString) Then
            Return True
        End If

        If inputString.Trim().Length = 0 Then
            Return True
        End If

        Return False
    End Function

    Private Sub ParseData()
        Dim dbFormat As cDatabaseFormat
        Dim fileName As String = ""
        Dim tableName As String = ""
        Dim groupingVariable As String = ""
        Dim groupingValue As Object = vbNull
        pd.GetDataFile(1, dbFormat, fileName, tableName, groupingVariable, groupingValue)

        CheckMeansAndIntercepts()

        Dim i As Integer = 0
        Dim unfilteredVariables As New ArrayList

        'Parse Information For Factor + Variable Data
        pd.GetDataFile(1, 18, fileName, tableName, groupingVariable, groupingValue)
        Dim SEM As New AmosEngine
        'SEM.ModelMeansAndIntercepts()
        'SEM.NonPositive()
        SEM.BeginGroup(fileName)
        For i = 1 To SEM.DataFileNVariables
            If (SEM.InputVariableIsNumeric(i)) And Not (SEM.InputVariableHasMissingValues(i)) Then
                unfilteredVariables.Add(SEM.InputVariableName(i))
                SEM.AStructure(SEM.InputVariableName(i) & " ()")
            End If
        Next
        SEM.Dispose()


        'Filter Stuff
        Dim frmVariables As New frmVariables
        Dim factor As String 'KEY FOR DICTIONARY (3 LETTER NICKNAME)
        Dim factorVariables As ArrayList 'LIST OF FULL-NAME VARIABLES
        filteredVariables = frmVariables.loadVariables(unfilteredVariables)

        ' Check To See If Variables Werse Selected
        If filteredVariables.Count = 0 Then
            MsgBox("Sorry, you did not select any variables. Terminating plugin.")
            Exit Sub
        End If

        'Establish Dictionary Pairs
        factors = New Dictionary(Of String, ArrayList)
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

        ClearCanvas()

        CreateUnobservedVarsFromFactors()
        CreateObservedVars()

        resizeObserved()
        drawCovariances()

        touchUpAll()

        linkData()
        ResetMeansAndIntercepts()
        pd.EditSelect()

    End Sub


    Private Sub CheckMeansAndIntercepts()
        estimatesChecked = pd.GetCheckBox("AnalysisPropertiesForm", "MeansInterceptsCheck").Checked
        pd.GetCheckBox("AnalysisPropertiesForm", "MeansInterceptsCheck").Checked = True
    End Sub

    Private Sub ResetMeansAndIntercepts()
        pd.GetCheckBox("AnalysisPropertiesForm", "MeansInterceptsCheck").Checked = estimatesChecked
    End Sub

    Private Sub ClearCanvas()
        While pd.PDElements.Count > 0
            pd.EditErase(pd.PDElements(1))
        End While
    End Sub

    Private Sub CreateUnobservedVarsFromFactors()
        verticalSeparation = (pageHeight - 72) / (filteredVariables.Count + factors.Count - 1)
        fontSize = verticalSeparation / 2
        If fontSize > 20 Then
            fontSize = 20
        End If

        Dim horizontalPosition As Double = pageWidth * 0.5

        ' Start the spacing half an inch below the top of the page
        Dim vPos As Double = 72 / 2

        For Each factor As KeyValuePair(Of String, ArrayList) In factors

            Dim factorName As String = factor.Key
            Dim factorVariables As ArrayList = factor.Value

            Dim verticalPosition As Double = (verticalSeparation * factorVariables.Count) / 2 + vPos
            Dim unobservedElement As PDElement = pd.DiagramDrawUnobserved(horizontalPosition, verticalPosition, 0.5, 0.7)

            unobservedElement.NameOrCaption = factorName
            'factors(Index).pdElement = unobservedElement
            unobservedElement.Width = fontSize * 4.5
            unobservedElement.Height = fontSize * 3

            unobservedElement.NameFontSize = CSng(fontSize)

            vPos = vPos + verticalSeparation * (factorVariables.Count + 1)

            pd.Refresh()
        Next
    End Sub

    Private Sub linkData()
        Dim SEM As AmosEngine = New AmosEngine
        pd.SpecifyModel(SEM)
        SEM.FitModel()
        SEM.Dispose()
    End Sub

    Private Function inchesToPoints(inches As Double) As Double
        Return inches * 72
    End Function

    Private Function pointsToInches(points As Double) As Double
        Return points / 72
    End Function

    Private Sub CreateObservedVars()
        Dim horizontalPosition As Double = pageWidth * 0.25
        Dim errorCounter As Integer = 1
        Dim errorhPos As Double = pageWidth * (0.1 + ((1 - ((fontSize * 6) / 100)) * 0.1))


        For Each factorDetail As KeyValuePair(Of String, ArrayList) In factors

            ' Set up the positioning for the factors
            Dim factorName As String = factorDetail.Key
            Dim factorVariables As ArrayList = factorDetail.Value
            Dim factorElement = pd.PDE(factorName)
            Dim factorPosition As Double = factorElement.originY
            Dim startVerticalPosition As Double = factorPosition - ((factorVariables.Count - 1) * verticalSeparation) / 2


            ' Loop through each item in the current factor
            Dim i As Integer = 0

            For Each factorVariable As String In factorVariables

                ' Create Variable Element
                Dim verticalPosition As Double = startVerticalPosition + (i * verticalSeparation)
                Dim observedElement As PDElement = pd.DiagramDrawObserved(horizontalPosition, verticalPosition, 0.2, 0.1)
                observedElement.Height = fontSize
                observedElement.Width = fontSize * 4

                observedElement.NameOrCaption = factorVariable
                observedElement.NameFontSize = CSng(fontSize)
                'currentFactor.linkedItems(itemIndex).pdElement = observedElement

                ' Create Error Element
                Dim errorElement As PDElement = pd.DiagramDrawUnobserved(errorhPos, verticalPosition, fontSize * 1.75, fontSize * 1.75)
                Dim errorPath As PDElement = pd.DiagramDrawPath(errorElement, observedElement)
                errorPath.Value1 = CType(1, String)
                errorElement.NameFontSize = CSng(fontSize)
                errorElement.NameOrCaption = "e" & errorCounter
                errorCounter = errorCounter + 1

                ' Add Path
                Dim path As PDElement = pd.DiagramDrawPath(factorElement, observedElement)
                ' If this is the first item, add a regression weight of 1
                If i = 0 Then
                    path.Value1 = CType(1, String)
                End If

                i = i + 1
            Next
        Next

        pd.Refresh()
    End Sub

    Private Function IsFactorRow(checkRow As Array) As Boolean
        ' Check if the row has variables in every row (except the first)
        Dim arraySize As Integer = checkRow.Length - 1
        For index As Integer = 0 To arraySize
            Dim currentRow As String = CType(checkRow.GetValue(index), String)
            If index = 0 And IsNullOrWhiteSpace(currentRow) = False Then
                Return False
            End If
            If index <> 0 And IsNullOrWhiteSpace(currentRow) Then
                Return False
            End If
        Next index

        Return True
    End Function

    Private Sub touchUpAll()
        Dim element As PDElement
        For Each element In pd.PDElements
            pd.EditTouchUp(element)
        Next
    End Sub

    Private Sub drawCovariances()
        Dim factorList As New ArrayList

        For Each factorDetail As KeyValuePair(Of String, ArrayList) In factors
            factorList.Add(factorDetail.Key)
        Next

        Dim currentFactor As String
        Dim linkFactor As String
        Dim currentFactorElement As PDElement
        Dim linkFactorElement As PDElement

        For i As Integer = 0 To factorList.Count - 1
            currentFactor = factorList(i)
            currentFactorElement = pd.PDE(currentFactor)

            For j As Integer = i + 1 To factors.Count - 1
                linkFactor = factorList(j)
                linkFactorElement = pd.PDE(linkFactor)
                pd.DiagramDrawCovariance()
                pd.DragMouse(currentFactorElement, linkFactorElement.originX, linkFactorElement.originY)
                pd.DiagramDrawCovariance()
            Next

        Next
    End Sub

    Private Sub resizeObserved()
        Dim x As PDElement
        Dim LargestWidth As Single, LargestHeight As Single
        LargestWidth = 0
        LargestHeight = 0
        For Each x In pd.PDElements
            If x.IsObservedVariable Then
                If x.NameWidth > LargestWidth Then LargestWidth = CSng(x.NameWidth)
                If x.NameHeight > LargestHeight Then LargestHeight = CSng(x.NameHeight)
            End If
        Next
        LargestWidth = CSng(LargestWidth * 1.4)
        LargestHeight = CSng(LargestHeight * 1.4)
        pd.UndoToHere()
        If LargestWidth > 0.2 And LargestHeight > 0.1 Then
            For Each x In pd.PDElements
                If x.IsObservedVariable Then
                    x.Width = LargestWidth
                    x.Height = LargestHeight
                End If
            Next
            pd.Refresh()
        End If
        pd.DiagramRedrawDiagram()
        pd.UndoResume()
    End Sub

End Class
