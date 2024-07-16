' Copyright 2015 Howard J Rudd
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, this software
' is distributed on an "AS IS" BASIS WITHOUT WARRANTIES OR CONDITIONS OF
' ANY KIND, either express or implied, not even for MERCHANTABILITY or
' FITNESS FOR A PARTICULAR PURPOSE. See the License for the specific language
' governing permissions and limitations under the License. You are free to use
' this code as you wish within the provisions of the license but it is your
' responsibility to test it and ensure it is fit for the use to which you
' intend to put it.
'
' _____________________________________________________________________________


' This class generates objects that contain sample matrices, other
' parameters and methods relating to subsets of the variables


'****************************************************************************************************************************************
'*  TOP LEVEL USER INPUTS MODULE
'*****************************************************************************************************************************************

Public Sub Main()

Dim start As Double, elatim As Double, elatim1 As Double, elatim2 As Double
start = Timer

' GENERAL DECLARATIONS

Dim All_Input_Variables As Collection, OutputVariables As Collection
Set All_Input_Variables = New Collection
Set OutputVariables = New Collection

Dim i As Long, j As Long, k As Long, m As Long, n As Long, U() As Double
n = 1000
ReDim U(1 To n)

' INPUT VARIABLES, SUBSET A ===============================================================================================================

Dim Input_Variable_SubSet_A As Class_Random_Variables_By_Correlation_Group
Set Input_Variable_SubSet_A = New Class_Random_Variables_By_Correlation_Group
All_Input_Variables.Add Input_Variable_SubSet_A

Input_Variable_SubSet_A.SubsetName = "Subset A (independent)"
Input_Variable_SubSet_A.NumVars = 5
Input_Variable_SubSet_A.NumIters = n
Input_Variable_SubSet_A.Size

'1st input varible in A
k = 1
Input_Variable_SubSet_A.VariableName(k) = "Input 1"
Input_Variable_SubSet_A.VariableSheet(k) = "Model"
Input_Variable_SubSet_A.VariableRange(k) = "B2"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_A.OrderedSample(i, k) = WorksheetFunction.GammaInv(U(i), 2.032, 4.117)
Next i

'2nd input varible in A
k = 2
Input_Variable_SubSet_A.VariableName(k) = "Input 2"
Input_Variable_SubSet_A.VariableSheet(k) = "Model"
Input_Variable_SubSet_A.VariableRange(k) = "B3"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_A.OrderedSample(i, k) = WorksheetFunction.NormInv(U(i), -1.3, 2.1)
Next i

'3rd input varible in A
k = 3
Input_Variable_SubSet_A.VariableName(k) = "Input 3"
Input_Variable_SubSet_A.VariableSheet(k) = "Model"
Input_Variable_SubSet_A.VariableRange(k) = "B4"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_A.OrderedSample(i, k) = TriangularInv(U(i), -1.1, 1.9, 5.2)
Next i

'4th input varible in A
k = 4
Input_Variable_SubSet_A.VariableName(k) = "Input 4"
Input_Variable_SubSet_A.VariableSheet(k) = "Model"
Input_Variable_SubSet_A.VariableRange(k) = "B5"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_A.OrderedSample(i, k) = UniformInv(U(i), 1.2, 5.721)
Next i

'5th input varible in A
k = 5
Input_Variable_SubSet_A.VariableName(k) = "Input 5"
Input_Variable_SubSet_A.VariableSheet(k) = "Model"
Input_Variable_SubSet_A.VariableRange(k) = "B6"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_A.OrderedSample(i, k) = WorksheetFunction.LogInv(U(i), 0.1, 0.2)
Next i

Input_Variable_SubSet_A.GenerateIndependentSample

' INPUT VARIABLES, SUBSET B ==================================================================================================================

Dim Input_Variable_SubSet_B As Class_Random_Variables_By_Correlation_Group
Set Input_Variable_SubSet_B = New Class_Random_Variables_By_Correlation_Group
All_Input_Variables.Add Input_Variable_SubSet_B

Input_Variable_SubSet_B.SubsetName = "Subset B (correlated)"
Input_Variable_SubSet_B.NumVars = 5
Input_Variable_SubSet_B.NumIters = n
Input_Variable_SubSet_B.Size

'First input varible in B
k = 1
Input_Variable_SubSet_B.VariableName(k) = "Input 6"
Input_Variable_SubSet_B.VariableSheet(k) = "Model"
Input_Variable_SubSet_B.VariableRange(k) = "B8"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_B.OrderedSample(i, k) = TriangularInv(U(i), 0.01, 1.1, 7.21)
Next i

'2nd input varible in B
k = 2
Input_Variable_SubSet_B.VariableName(k) = "Input 7"
Input_Variable_SubSet_B.VariableSheet(k) = "Model"
Input_Variable_SubSet_B.VariableRange(k) = "B9"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_B.OrderedSample(i, k) = WorksheetFunction.NormInv(U(i), 12.1, 5.2)
Next i

'3rd input varible in B
k = 3
Input_Variable_SubSet_B.VariableName(k) = "Input 8"
Input_Variable_SubSet_B.VariableSheet(k) = "Model"
Input_Variable_SubSet_B.VariableRange(k) = "B10"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_B.OrderedSample(i, k) = UniformInv(U(i), 4, 9.47)
Next i

'4th input varible in B
k = 4
Input_Variable_SubSet_B.VariableName(k) = "Input 9"
Input_Variable_SubSet_B.VariableSheet(k) = "Model"
Input_Variable_SubSet_B.VariableRange(k) = "B11"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_B.OrderedSample(i, k) = WorksheetFunction.LogInv(U(i), 1.03, 0.721)
Next i

'5th input varible in B
k = 5
Input_Variable_SubSet_B.VariableName(k) = "Input 10"
Input_Variable_SubSet_B.VariableSheet(k) = "Model"
Input_Variable_SubSet_B.VariableRange(k) = "B12"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_B.OrderedSample(i, k) = TriangularInv(U(i), -2, 5.7, 15.721)
Next i

Input_Variable_SubSet_B.CorrelationMatrixSheet = "Correlation Matrix 1"
Input_Variable_SubSet_B.CorrelationMatrixRange = "A1:E5"
Input_Variable_SubSet_B.GenerateCorrelatedSample

' INPUT VARIABLES, SUBSET C

Dim Input_Variable_SubSet_C As Class_Random_Variables_By_Correlation_Group
Set Input_Variable_SubSet_C = New Class_Random_Variables_By_Correlation_Group
All_Input_Variables.Add Input_Variable_SubSet_C

Input_Variable_SubSet_C.SubsetName = "Subset C (correlated)"
Input_Variable_SubSet_C.NumVars = 5
Input_Variable_SubSet_C.NumIters = n
Input_Variable_SubSet_C.Size

'First input varible in C
k = 1
Input_Variable_SubSet_C.VariableName(k) = "Input 11"
Input_Variable_SubSet_C.VariableSheet(k) = "Model"
Input_Variable_SubSet_C.VariableRange(k) = "B14"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_C.OrderedSample(i, k) = UniformInv(U(i), -1, 12)
Next i

'2nd input varible in C
k = 2
Input_Variable_SubSet_C.VariableName(k) = "Input 12"
Input_Variable_SubSet_C.VariableSheet(k) = "Model"
Input_Variable_SubSet_C.VariableRange(k) = "B15"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_C.OrderedSample(i, k) = TriangularInv(U(i), 2.1, 3.92, 15.721)
Next i

'3rd input varible C
k = 3
Input_Variable_SubSet_C.VariableName(k) = "Input 12a"
Input_Variable_SubSet_C.VariableSheet(k) = "Model"
Input_Variable_SubSet_C.VariableRange(k) = "B16"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_C.OrderedSample(i, k) = WorksheetFunction.NormInv(U(i), 37.8, 3.26)
Next i

'4th input varible C
k = 4
Input_Variable_SubSet_C.VariableName(k) = "Input 14"
Input_Variable_SubSet_C.VariableSheet(k) = "Model"
Input_Variable_SubSet_C.VariableRange(k) = "B17"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_C.OrderedSample(i, k) = WorksheetFunction.GammaInv(U(i), 2.1, 5.721)
Next i

'5th input varible C
k = 5
Input_Variable_SubSet_C.VariableName(k) = "Input 15"
Input_Variable_SubSet_C.VariableSheet(k) = "Model"
Input_Variable_SubSet_C.VariableRange(k) = "B18"
U = unifun(n)
For i = 1 To n
    Input_Variable_SubSet_C.OrderedSample(i, k) = TriangularInv(U(i), 0.01, 2.92, 5.721)
Next i

Input_Variable_SubSet_C.CorrelationMatrixSheet = "Correlation Matrix 2"
Input_Variable_SubSet_C.CorrelationMatrixRange = "A2:E6"
Input_Variable_SubSet_C.GenerateCorrelatedSample

' OUTPUT VARIABLES

Dim Output_Variable As Class_Random_Variables_By_Correlation_Group
Set Output_Variable = New Class_Random_Variables_By_Correlation_Group
OutputVariables.Add Output_Variable

Output_Variable.NumVars = 2
Output_Variable.NumIters = n
Output_Variable.Size

' First output variable

Output_Variable.VariableName(1) = "Output 1"
Output_Variable.VariableSheet(1) = "Model"
Output_Variable.VariableRange(1) = "B21"

' Second output variable

Output_Variable.VariableName(2) = "Output 2"
Output_Variable.VariableSheet(2) = "Model"
Output_Variable.VariableRange(2) = "B22"

' END OF USER INPUTS

elatim1 = Timer - start

Call RunModel(All_Input_Variables, OutputVariables)
Call Graphs(All_Input_Variables, NumBins:=20, NumPoints:=100, SheetTitle:="Input variables")
Call Graphs(OutputVariables, NumBins:=20, NumPoints:=100, SheetTitle:="Output variables")

elatim2 = Timer - start - elatim1

elatim = elatim1 + elatim2

MsgBox "Number of iterations = " & n & vbNewLine & _
       "Time to generate input sample = " & elatim1 & " seconds" & vbNewLine & _
       "Time to run spreadsheet model = " & elatim2 & " seconds" & vbNewLine & _
       "Total time = " & elatim & " seconds."
       
End Sub



'****************************************************************************************************************************************
'*  ADMIN FUNCTIONS
'*****************************************************************************************************************************************


Public Function RunModel(All_Input_Variables As Collection, OutputVariables As Collection)

Dim Subset As Class_Random_Variables_By_Correlation_Group
Dim Subset2 As Class_Random_Variables_By_Correlation_Group
Dim num As Long, i As Long, j As Long

' Check that all subsets have the same value of NumIters
For Each Subset In All_Input_Variables
    num = Subset.NumIters
    For Each Subset2 In All_Input_Variables
        If Not (num = Subset2.NumIters) Then
            Err.Raise Number:=vbObjectError + 5, _
            Source:="RunModel", _
            Description:="Input variable subsets don't all have the same number of iterations"
            Exit Function
        End If
    Next Subset2
Next Subset

For Each Subset In OutputVariables
    If Not (num = Subset.NumIters) Then
            Err.Raise Number:=vbObjectError + 6, _
            Source:="RunModel", _
            Description:="Output variable subsets don't all have the same number of iterations as the input variables"
            Exit Function
    End If
Next Subset

For Each Subset In All_Input_Variables
    For j = 1 To Subset.NumVars
        Subset.TempStore(j) = ThisWorkbook.Sheets(Subset.VariableSheet(j)).Range(Subset.VariableRange(j))
    Next j
Next Subset

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
For i = 1 To num
    For Each Subset In All_Input_Variables
        For j = 1 To Subset.NumVars
            ThisWorkbook.Sheets(Subset.VariableSheet(j)).Range(Subset.VariableRange(j)).Value2 = Subset.Sample(i, j)
        Next j
    Next Subset
    Application.Calculate
    For Each Subset2 In OutputVariables
        For j = 1 To Subset2.NumVars
            Subset2.Sample(i, j) = ThisWorkbook.Sheets(Subset2.VariableSheet(j)).Range(Subset2.VariableRange(j)).Value2
        Next j
    Next Subset2
Next i
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

For Each Subset In All_Input_Variables
    For j = 1 To Subset.NumVars
        ThisWorkbook.Sheets(Subset.VariableSheet(j)).Range(Subset.VariableRange(j)) = Subset.TempStore(j)
    Next j
Next Subset

For Each Subset In OutputVariables
    Subset.GenerateOrderedSample
Next Subset

End Function


Public Function IsSheetThere(ByVal SheetName As String) As Boolean

' Usage: Y = IsSheetThere(SheetName)
'
' Determines whether or not a sheet named "sheetName" exists in the active
' workbook.

Dim TestSheet As Worksheet
Dim bReturn As Boolean
bReturn = False
    For Each TestSheet In ThisWorkbook.Worksheets
        If TestSheet.Name = SheetName Then
            bReturn = True
            Exit For
        End If
    Next TestSheet
IsSheetThere = bReturn
End Function

Public Function TestAndAdd(ByVal SheetName As String, Optional Zoom As Integer) As String

' Usage: Y = TestAndAdd(String, Zoom)
'
' Finds the smallest value of i such that a sheet called sheetNamei doesn't
' already exist and then creates a sheet called "SheetNamei".
'
' That is, it looks in this workbook for a worksheet named sheetName. If it
' doesn't find it, it creates it. If it does find sheetName, it searches for
' sheetName1. If it doesn't find sheetName1 it creates it. If it does find it,
' it sesarches for sheetName2 and so on.
'
' Returns the name of the added sheet as a string.

Dim TestName As String
Dim TempName As String
Dim i As Integer

' Assign the desired name to the string variable testName
If IsSheetThere(SheetName) = False Then
    TestName = SheetName
Else
    i = 2
    TempName = SheetName & i
    Do While IsSheetThere(TempName) = True
        i = i + 1
        TempName = SheetName & i
    Loop
TestName = TempName
End If

'Add a new sheet and change its .Name property to testName
Dim newSheet As Worksheet
Set newSheet = ThisWorkbook.Sheets.Add
newSheet.Name = TestName

' Adjust magnification
If Not IsMissing(Zoom) Then
    ThisWorkbook.Sheets(TestName).Activate
    ActiveWindow.Zoom = Zoom
End If

' Return name of added sheet as string
TestAndAdd = TestName
    
End Function


Public Function Graphs(VarSet As Collection, NumBins As Long, NumPoints As Long, SheetTitle As String)

' Usage: Graphs(VarSet As Collection, NumBins As Long, NumPoints As Long, _
                SheetTitle As String)
'
' Creates a new worksheet containing an array of embedded charts showing
' the frequency distributions of the outputs of the spreadsheet model.
'
' Calculates statistics and writes them into the worksheet.
'
' VarSet is a collection of Class_Random_Variables_By_Correlation_Group objects.
'
' NumBins is the number of bins the the histogram.
'
' NumPoints is the number of points to be plotted in a cumulative distribution
' plot.
' _____________________________________________________________________________

' DECLARATIONS (in aphabetical order)

Dim AbsBinLeft As Double
Dim AbsBinRight As Double
Dim Absomax As Double
Dim BinContentCount() As Long
Dim BinLabels() As String
Dim BinLeft As Double
Dim BinRight As Double
Dim ChtHeightCells As Integer
Dim ChtHeightPts As Integer
Dim ChtHorizGapCells As Long
Dim ChtTopLeftCellColumn As Integer
Dim ChtTopLeftCellRow As Integer
Dim ChtVertGapCells As Long
Dim ChtWidthCells As Integer
Dim ChtWidthPts As Integer
Dim CumChtAbscissaRange() As Range
Dim CumChtOrdinateRange() As Range
Dim CumChtTitleRange() As Range
Dim CumSubsetRange As Range
Dim CumSubsetTitleRange As Range
Dim CumulativePlot() As Object
Dim DeltaX As Double
Dim Distribution() As Object
Dim ExponentX As Long
Dim EXsquared As Double
Dim Fstep As Long
Dim FXabscissa() As Double
Dim FXordinate() As Double
Dim HistChtAbscissaRange() As Range
Dim HistChtOrdinateRange() As Range
Dim HistChtTitleRange() As Range
Dim Histogram() As Object
Dim HistSigFigs As Integer
Dim HistSubsetRange As Range
Dim HistSubsetTitleRange As Range
Dim HOffset As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim LeftX As Double
Dim m As Long
Dim n As Long
Dim NumIters As Long
Dim NumVars As Long
Dim PreviousSubsetNumVars As Long
Dim PreviousSubsets As Long
Dim r As Long
Dim RightX As Double
Dim SampleMax As Double
Dim SampleMean As Double
Dim SampleMin As Double
Dim SampleSigma As Double
Dim SheetName As String
Dim StatsTopLeftCellColumn As Long
Dim StatsTopLeftCellRow As Long
Dim StatSubsetRange As Range
Dim str1 As String
Dim str2 As String
Dim StrExponentX As String
Dim StrX As String
Dim Subset As Class_Random_Variables_By_Correlation_Group
Dim SubsetRowStride As Long
Dim SumFormula As String
Dim TotalPreviousNumVars As Long
Dim VOffset As Long
Dim x As Double
Dim y As Long

SheetName = TestAndAdd(SheetTitle, 75)

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

ChtWidthCells = 5
ChtHeightCells = 15
ChtHorizGapCells = 2
ChtVertGapCells = 2

PreviousSubsets = 0
PreviousSubsetNumVars = 0
TotalPreviousNumVars = 0

For Each Subset In VarSet

    NumVars = Subset.NumVars
    NumIters = Subset.NumIters

    ReDim BinContentCount(1 To NumBins, 1 To NumVars)
    ReDim BinLabels(1 To NumBins, 1 To NumVars)
    ReDim FXordinate(1 To NumPoints, 1 To NumVars)
    ReDim FXabscissa(1 To NumPoints, 1 To NumVars)

    ReDim HistChtTitleRange(1 To NumVars)
    ReDim HistChtAbscissaRange(1 To NumVars)
    ReDim HistChtOrdinateRange(1 To NumVars)

    ReDim CumChtTitleRange(1 To NumVars)
    ReDim CumChtAbscissaRange(1 To NumVars)
    ReDim CumChtOrdinateRange(1 To NumVars)
    
    Fstep = Int(NumIters / NumPoints)
    
  ' Set gap between top of sheet and top of first graph of Subset
    SubsetRowStride = TotalPreviousNumVars * (ChtHeightCells + ChtVertGapCells) + ChtVertGapCells
    
  ' Set left column for both data tables
    HOffset = 2 * (ChtWidthCells + 1) + 10 + 2 * TotalPreviousNumVars
        
  ' Define worksheet ranges to contain histogram data
    Set HistSubsetRange = ThisWorkbook.Sheets(SheetName) _
        .Range(Cells(1, HOffset + 1), Cells(NumBins + 3, HOffset + 2 * NumVars))
    With HistSubsetRange
        Set HistSubsetTitleRange = .Range(Cells(1, 1), Cells(1, 2 * NumVars))
        For k = 1 To NumVars
            Set HistChtTitleRange(k) = .Range(Cells(2, 2 * k - 1), Cells(2, 2 * k))
            Set HistChtAbscissaRange(k) = .Range(Cells(3, 2 * k - 1), Cells(NumBins + 2, 2 * k - 1))
            Set HistChtOrdinateRange(k) = .Range(Cells(3, 2 * k), Cells(NumBins + 2, 2 * k))
        Next k
    End With
    
  ' Define worksheet ranges to contain cumulative chart data
    Set CumSubsetRange = ThisWorkbook.Sheets(SheetName) _
        .Range(Cells(NumBins + 5, HOffset + 1), Cells(NumBins + 4 + NumPoints + 3, HOffset + 2 * NumVars))
    With CumSubsetRange
        Set CumSubsetTitleRange = .Range(Cells(1, 1), Cells(1, 2 * NumVars))
        For k = 1 To NumVars
            Set CumChtTitleRange(k) = .Range(Cells(2, 2 * k - 1), Cells(2, 2 * k))
            Set CumChtAbscissaRange(k) = .Range(Cells(3, 2 * k - 1), Cells(NumPoints + 2, 2 * k - 1))
            Set CumChtOrdinateRange(k) = .Range(Cells(3, 2 * k), Cells(NumPoints + 2, 2 * k))
        Next k
    End With
       
  ' Translate chart dimensions into points
    With ThisWorkbook.Sheets(SheetName)
        ChtWidthPts = ChtWidthCells * (Cells(1, 2).left - Cells(1, 1).left)
        ChtHeightPts = ChtHeightCells * (Cells(2, 1).Top - Cells(1, 1).Top)
    End With
    ' Debug.Print "Chart width = " & ChtWidthPts & ", height = " & ChtHeightPts

  ' HISTOGRAMS
  
    HistSigFigs = 4
  
    For k = 1 To NumVars
        
    ' Calculate histogram data for the kth variable in Subset
    
      ' Calculate exponent for rounding
        With Subset
            SampleMin = .Min(k)
            SampleMax = .Max(k)
            SampleMean = .SampleMean(k)
            SampleSigma = .Variance(k)
        End With
        SampleSigma = Sqr(SampleSigma)
        
        Absomax = WorksheetFunction.Max(Abs(SampleMin), Abs(SampleMax))
        ExponentX = expon(Absomax)
        
      ' Calculate histogram bin dimensions
      
        LeftX = WorksheetFunction.Max(SampleMin, (SampleMean - 3 * SampleSigma))
        RightX = WorksheetFunction.Min(SampleMax, (SampleMean + 3 * SampleSigma))
        
        If LeftX < 0 Then
            LeftX = -WorksheetFunction.RoundUp(Abs(LeftX), (HistSigFigs - ExponentX - 1))
        Else
            LeftX = WorksheetFunction.RoundDown(LeftX, (HistSigFigs - ExponentX - 1))
        End If
       
        DeltaX = (RightX - LeftX) / NumBins
        DeltaX = WorksheetFunction.Round(DeltaX, (HistSigFigs - ExponentX - 1))
    
      ' Initialise histogram bins
        For i = 1 To NumBins
            BinContentCount(i, k) = 0
        Next i
    
      ' Populate histogram bins
        For i = 1 To NumIters
            x = Subset.OrderedSample(i, k)
            BinLeft = LeftX
            For j = 1 To NumBins
                BinRight = BinLeft + DeltaX
                If (x >= BinLeft) And (x < BinRight) Then
                    BinContentCount(j, k) = BinContentCount(j, k) + 1
                End If
                BinLeft = BinRight
            Next j
        Next i

      ' Generate bin label text
        BinLeft = LeftX
        For j = 1 To NumBins
            BinRight = BinLeft + DeltaX
            BinRight = WorksheetFunction.Round(BinRight, (HistSigFigs - 1 - ExponentX))
            If Not (Abs(DeltaX) < 10 ^ HistSigFigs And Abs(DeltaX) > 0.1) Then
                BinLabels(j, k) = Format(BinLeft, "Scientific") & " to " & Format(BinRight, "Scientific")
            Else
                BinLabels(j, k) = BinLeft & " to " & BinRight
            End If
            BinLeft = BinRight
        Next j

    ReDim Histogram(1 To NumVars)
                
   ' Populate histogram worksheet ranges
      ' Subset title
        With HistSubsetTitleRange
            .Merge
            .HorizontalAlignment = xlCenter
            .Value = "Histogram data, " & Subset.SubsetName
        End With
        
      ' Dataset title
        With HistChtTitleRange(k)
            .Merge
            .HorizontalAlignment = xlCenter
            .Value = Subset.VariableName(k)
        End With
        
      ' Abscissa data
        With HistChtAbscissaRange(k)
            For i = 1 To NumBins
                .Cells(i, 1) = BinLabels(i, k)
            Next i
            str1 = .Address(ReferenceStyle:=xlR1C1)
            SumFormula = "=SUM(" & str1 & ")"
            .Cells(NumBins + 1, 1).Formula = SumFormula
        End With
        
      ' Ordinate data
        With HistChtOrdinateRange(k)
            For i = 1 To NumBins
                .Cells(i, 1) = BinContentCount(i, k)
            Next i
            str1 = .Address(ReferenceStyle:=xlR1C1)
            SumFormula = "=SUM(" & str1 & ")"
            .Cells(NumBins + 1, 1).Formula = SumFormula
        End With
        
    Next k
    
  ' Create histograms
    For i = 1 To NumVars
        ChtTopLeftCellColumn = ChtHorizGapCells
        ChtTopLeftCellRow = (i - 1) * (ChtHeightCells + ChtVertGapCells) + SubsetRowStride + 1
        Set Histogram(i) = ThisWorkbook.Sheets(SheetName).ChartObjects.Add( _
            left:=Cells(ChtTopLeftCellRow, ChtTopLeftCellColumn).left, _
            Top:=Cells(ChtTopLeftCellRow, ChtTopLeftCellColumn).Top, _
            Width:=ChtWidthPts, _
            Height:=ChtHeightPts)
        With Histogram(i).chart
            .SetSourceData Source:=HistChtOrdinateRange(i)
            .SeriesCollection(1).XValues = HistChtAbscissaRange(i)
            .ChartType = 51
            .ChartGroups(1).GapWidth = 0
            .SetElement (msoElementLegendNone)
            .HasTitle = True
            .chartTitle.text = Subset.VariableName(i)
            .chartTitle.Font.FontStyle = "Regular"
            .chartTitle.Font.Size = 10
            .Axes(xlValue).TickLabelPosition = xlNone
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlCategory).TickLabels.Orientation = xlUpward
            .ChartArea.Border.LineStyle = xlNone
        End With
    Next i
        
  ' CUMULATIVE DISTRIBUTION PLOTS

    For k = 1 To NumVars
    
  ' Generate cumulative distribution data for the kth variable in Subset
    
        FXabscissa(1, k) = Subset.CumulativeAbscissa(1, k)
        FXordinate(1, k) = Subset.CumulativeOrdinate(1, k)
        For i = 2 To NumPoints - 1
            FXabscissa(i, k) = Subset.CumulativeAbscissa((i - 1) * Fstep, k)
            FXordinate(i, k) = Subset.CumulativeOrdinate((i - 1) * Fstep, k)
        Next i
        FXabscissa(NumPoints, k) = Subset.CumulativeAbscissa(NumIters, k)
        FXordinate(NumPoints, k) = Subset.CumulativeOrdinate(NumIters, k)
        
   ' Populate Cumulative distribution worksheet ranges
      ' Subset title
        With CumSubsetTitleRange
            .Merge
            .HorizontalAlignment = xlCenter
            .Value = "Cumulative distribution data, " & Subset.SubsetName
        End With
        
      ' Dataset title
        With CumChtTitleRange(k)
            .Merge
            .HorizontalAlignment = xlCenter
            .Value = Subset.VariableName(k)
        End With
        
      ' Abscissa data
        With CumChtAbscissaRange(k)
            For i = 1 To NumPoints
                .Cells(i, 1) = FXabscissa(i, k)
            Next i
            str1 = .Address(ReferenceStyle:=xlR1C1)
            SumFormula = "=SUM(" & str1 & ")"
            .Cells(NumPoints + 1, 1).Formula = SumFormula
        End With
        
      ' Ordinate data
        With CumChtOrdinateRange(k)
            For i = 1 To NumPoints
                .Cells(i, 1) = FXordinate(i, k)
            Next i
            str1 = .Address(ReferenceStyle:=xlR1C1)
            SumFormula = "=SUM(" & str1 & ")"
            .Cells(NumPoints + 1, 1).Formula = SumFormula
        End With
        
    Next k
    
    ReDim CumulativePlot(1 To NumVars)
        
  ' Create cumulative distribution charts
    For i = 1 To NumVars
        ChtTopLeftCellColumn = (ChtWidthCells + 1) + 2
        ChtTopLeftCellRow = (i - 1) * (ChtHeightCells + ChtVertGapCells) + SubsetRowStride + 1
        Set CumulativePlot(i) = ThisWorkbook.Sheets(SheetName).ChartObjects.Add( _
            left:=Cells(ChtTopLeftCellRow, ChtTopLeftCellColumn).left, _
            Top:=Cells(ChtTopLeftCellRow, ChtTopLeftCellColumn).Top, _
            Width:=ChtWidthPts, _
            Height:=ChtHeightPts)
        With CumulativePlot(i).chart
            .SetSourceData Source:=CumChtOrdinateRange(i)
            .ChartType = xlXYScatter
            With .SeriesCollection(1)
                .XValues = CumChtAbscissaRange(i)
                .MarkerStyle = xlNone
                .Border.LineStyle = xlContinuous
            End With
            .ChartGroups(1).GapWidth = 0
            .SetElement (msoElementLegendNone)
            .HasTitle = True
            .chartTitle.text = Subset.VariableName(i)
            .chartTitle.Font.FontStyle = "Regular"
            .chartTitle.Font.Size = 10
            .Axes(xlValue).TickLabelPosition = xlTickLabelPositionNone
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 1
            .Axes(xlCategory).MinimumScale = Subset.Min(i)
            .Axes(xlCategory).MaximumScale = Subset.Max(i)
            .Axes(xlCategory).TickLabels.Orientation = xlUpward
            .Axes(xlCategory).TickLabels.NumberFormat = "General"
            .ChartArea.Border.LineStyle = xlNone
        End With
    Next i
    
  ' Write statistics into worksheet

    For k = 1 To NumVars
    
      ' Define worksheet ranges to contain statistics
        StatsTopLeftCellRow = (k - 1) * (ChtHeightCells + ChtVertGapCells) + SubsetRowStride + 1
        StatsTopLeftCellColumn = 2 * ChtWidthCells + 4
        Set StatSubsetRange = ThisWorkbook.Sheets(SheetName).Cells(StatsTopLeftCellRow, StatsTopLeftCellColumn)

        With StatSubsetRange
        
                With .Range(Cells(1, 1), Cells(1, 2))
                        .Merge
                        .HorizontalAlignment = xlCenter
                        .Value2 = Subset.VariableName(k) & " Statistics"
                End With
                        
                .Cells(2, 1).Value2 = "Min"
                .Cells(3, 1).Value2 = "10th percentile"
                .Cells(4, 1).Value2 = "Lower quartile"
                .Cells(5, 1).Value2 = "Median"
                .Cells(6, 1).Value2 = "Mean"
                .Cells(7, 1).Value2 = "Upper quartile"
                .Cells(8, 1).Value2 = "90th percentile"
                .Cells(9, 1).Value2 = "Max"
                .Cells(10, 1).Value2 = "Variance"
                .Cells(11, 1).Value2 = "St. dev."

                .Cells(2, 2).Value2 = Subset.Min(k)
                .Cells(3, 2).Value2 = Subset.Quantile(0.1, k)
                .Cells(4, 2).Value2 = Subset.Quantile(0.25, k)
                .Cells(5, 2).Value2 = Subset.Quantile(0.5, k)
                .Cells(6, 2).Value2 = Subset.SampleMean(k)
                .Cells(7, 2).Value2 = Subset.Quantile(0.75, k)
                .Cells(8, 2).Value2 = Subset.Quantile(0.9, k)
                .Cells(9, 2).Value2 = Subset.Max(k)
                .Cells(10, 2).Value2 = Subset.Variance(k)
                .Cells(11, 2).Value2 = Sqr(Subset.Variance(k))

                .Cells(1, 1).EntireColumn.AutoFit
        End With
    
    Next k
    
    PreviousSubsets = PreviousSubsets + 1
    PreviousSubsetNumVars = NumVars
    TotalPreviousNumVars = TotalPreviousNumVars + NumVars
    
    Application.Calculate

Next Subset

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Function


'****************************************************************************************************************************************
'*  Generarate Simulated Numbers
'****************************************************************************************************************************************



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Random_Variables_By_Correlation_Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text



' This class generates objects that contain sample matrices, other
' parameters and methods relating to subsets of the variables

Private i As Long
Private j As Long

Private pNumVars As Long
Private pNumIters As Long

Private pSubsetName As String
Private pVariableName() As String
Private pVariableSheet() As String
Private pVariableRange() As String
Private v As Double
Private w As Double
Private x As Double
Private pTempStore() As Variant

Private pCorrelationMatrixSheet As String
Private pCorrelationMatrixRange As String
Private pCorrelationMatrix() As Double
Private pGap As Long

Private pQuantileIndex As Long
Private pQuantileAbscissa As Double
Private pVariance As Double

Private pOrderedSample() As Double
Private pTempSample() As Double
Private pSample() As Double

Public Property Let SubsetName(text As String)
    pSubsetName = text
End Property

Public Property Get SubsetName() As String
     SubsetName = pSubsetName
End Property

Public Property Let NumVars(Number As Long)
    pNumVars = Number
End Property

Public Property Get NumVars() As Long
    NumVars = pNumVars
End Property

Public Property Let NumIters(Number As Long)
    pNumIters = Number
End Property

Public Property Get NumIters() As Long
    NumIters = pNumIters
End Property

Public Property Let VariableName(Index As Long, Desc As String)
    pVariableName(Index) = Desc
End Property

Public Property Get VariableName(Index As Long) As String
    VariableName = pVariableName(Index)
End Property

Public Property Let VariableSheet(Index As Long, Desc As String)
    pVariableSheet(Index) = Desc
End Property

Public Property Get VariableSheet(Index As Long) As String
    VariableSheet = pVariableSheet(Index)
End Property

Public Property Let VariableRange(Index As Long, Desc As String)
    pVariableRange(Index) = Desc
End Property

Public Property Get VariableRange(Index As Long) As String
    VariableRange = pVariableRange(Index)
End Property

Public Property Let TempStore(Index As Long, Value As Variant)
    pTempStore(Index) = Value
End Property

Public Property Get TempStore(Index As Long) As Variant
    TempStore = pTempStore(Index)
End Property

Public Property Let OrderedSample(rowIndex As Long, colIndex As Long, Value As Double)
    pOrderedSample(rowIndex, colIndex) = Value
End Property

Public Property Get OrderedSample(rowIndex As Long, colIndex As Long) As Double
    OrderedSample = pOrderedSample(rowIndex, colIndex)
End Property

Public Property Let Sample(rowIndex As Long, colIndex As Long, Value As Double)
    pSample(rowIndex, colIndex) = Value
End Property

Public Property Get Sample(rowIndex As Long, colIndex As Long) As Double
    Sample = pSample(rowIndex, colIndex)
End Property

Public Property Get CorrelationMatrix(rowIndex As Long, colIndex As Long) As Double
    CorrelationMatrix = pCorrelationMatrix(rowIndex, colIndex)
End Property

Public Property Let CorrelationMatrixSheet(Desc As String)
    pCorrelationMatrixSheet = Desc
End Property

Public Property Get CorrelationMatrixSheet() As String
    CorrelationMatrixSheet = pCorrelationMatrixSheet
End Property

Public Property Let CorrelationMatrixRange(Desc As String)
    pCorrelationMatrixRange = Desc
End Property

Public Property Get CorrelationMatrixRange() As String
    CorrelationMatrixRange = pCorrelationMatrixRange
End Property

Public Sub Size()

    ReDim pVariableName(1 To pNumVars)
    ReDim pVariableSheet(1 To pNumVars)
    ReDim pVariableRange(1 To pNumVars)
    ReDim pTempStore(1 To pNumVars)
    ReDim pTempSample(1 To pNumIters) As Double
    ReDim pOrderedSample(1 To pNumIters, 1 To pNumVars)
    ReDim pSample(1 To pNumIters, 1 To pNumVars)
    
End Sub

Public Sub GenerateIndependentSample()

' This routine makes a copy of the ordered sample matrix and randomly shuffles
' the elements of each column. Each column is shuffled independently of the
' others.
       
    pSample = pOrderedSample

    For i = 1 To pNumVars
        Call shuffle(pSample, i)
    Next i
    
End Sub

Private Sub ImportCorrelationMatrix()

' Imports values from a worksheet and assigns them to an array inside the
' object. The values must be contained in the upper triangular half of a square
' array of cells. These are assumed to form the upper triangular half of a
' symmetric matrix. The lower half is generated by reflection in the diagonal.
' The user only needs to enter the upper triangular half. Row and column
' headings are added to the user-inputted array to enable the user to check
' that the correlation coefficients s/he entered correspond to the correct pair
' of variables. The colour of the cells from which values were imported is
' changed to enable the user to tell that the correct cells have been imported.
' A copy of the full array is printed below the user-inputted array to confirm
' that it has been imported correctly.

    ReDim pCorrelationMatrix(1 To pNumVars, 1 To pNumVars)
            
    With ThisWorkbook.Sheets(pCorrelationMatrixSheet).Range(pCorrelationMatrixRange)
        
      ' Test that the chosen range is square
        If Not .Rows.Count = .Columns.Count Then
      ' Error message
            Err.Raise Number:=vbObjectError + 1, _
            Source:="ImportCorrelationMatrix", _
            Description:="Correlation matrix range not square"
        End If
        
      ' Test that the range is the correct size for the number of variables
        If Not (.Rows.Count = pNumVars And .Columns.Count = pNumVars) Then
      ' Error message
            Err.Raise Number:=vbObjectError + 2, _
            Source:="ImportCorrelationMatrix", _
            Description:="Correlation matrix size doesn't match number of variables"
        End If
        
      ' Test that the diagonal elements are all unity
        For i = 1 To pNumVars
        ' Debug.Print "diagonal element = " & .Cells(i, i).Value
            If Not (.Cells(i, i).Value = 1) Then
                Debug.Print "error criterion met: " & i
                Err.Raise Number:=vbObjectError + 3, _
                Source:="ImportCorrelationMatrix", _
                Description:="Correlation matrix diagonal elements not all unity"
            End If
        Next i
        
        For i = 1 To pNumVars
            For j = i To pNumVars
              
              ' Import cells into array
                pCorrelationMatrix(i, j) = .Cells(i, j)
                
              ' Write column headings above or below matrix range
                If .Cells(1, 1).Row > 1 Then
                    'Debug.Print .Cells(1, 1).Row
                   .Cells(1, j).Offset(rowOffset:=-1).Value = pVariableName(j)
                Else
                   .Cells(1 + pNumVars, j).Value = pVariableName(j)
                End If
                
              ' Write row headinsg to the right of matrix range
               .Cells(i, pNumVars + 1).Value = pVariableName(i)
              
              ' Shade cells or, if already shaded, change colour
                With .Cells(i, j).Interior
                    If .Color = RGB(240, 240, 240) Then
                        .Color = RGB(200, 240, 240)
                    Else
                        .Color = RGB(240, 240, 240)
                    End If
                End With
            Next j
        Next i
    End With
        
    For i = 2 To pNumVars
        For j = 1 To i - 1
            pCorrelationMatrix(i, j) = pCorrelationMatrix(j, i)
        Next j
    Next i
    
    ' Write a copy of the full correlation matrix below the user-entered one to
    ' confirm that it has been imported correctly.
    pGap = 5
    With ThisWorkbook.Sheets(pCorrelationMatrixSheet).Range(pCorrelationMatrixRange)
        For i = 1 To pNumVars
            For j = 1 To pNumVars
                .Cells(i + pNumVars + pGap, j).Value = pCorrelationMatrix(i, j)
                .Cells(i + pNumVars + pGap, pNumVars + 1).Value = pVariableName(i)
                .Cells(pNumVars + pGap, j).Value = pVariableName(j)
            Next j
        Next i
    End With
    
End Sub

Public Sub GenerateCorrelatedSample()
    Call ImportCorrelationMatrix
    pSample = ic(pOrderedSample, pCorrelationMatrix)
End Sub

Public Sub GenerateOrderedSample()
' need some tests to test whether pOrderedSample already contains numbers.
pOrderedSample = pSample
    For j = 1 To pNumVars
        Call quicksort(pOrderedSample, j)
    Next j
End Sub

Public Property Get Min(Index As Long) As Double
    Min = pOrderedSample(1, Index)
End Property

Public Property Get Max(Index As Long) As Double
    Max = pOrderedSample(pNumIters, Index)
End Property

Public Property Get Quantile(Probability As Double, colIndex As Long) As Double

' Returns a value, x, such that the fraction of members of the sample that are
' <= x is equal to "probability". Values that are not exactly equal to numbers
' in the sample are calculated by linear interpolation between the two adjacent
' sample elements. "colIndex" is the column number.
'

' Test that probability is valid
    If (Probability < 1 / (NumIters + 1)) Or (Probability > NumIters / (NumIters + 1)) Then
      ' Error message
        Err.Raise Number:=vbObjectError + 4, _
        Source:="Subset.Quantile", _
        Description:="Probability outside of valid range"
        Exit Function
    End If
    
    pQuantileAbscissa = Probability * (pNumIters + 1)
    pQuantileIndex = Int(pQuantileAbscissa)
    Quantile = pOrderedSample(pQuantileIndex, colIndex) + _
              (pOrderedSample((pQuantileIndex + 1), colIndex) - _
               pOrderedSample(pQuantileIndex, colIndex)) * _
              (pQuantileAbscissa - pQuantileIndex)

End Property

Private Function pSampleSum(colIndex As Long) As Double
    x = 0
    For i = 1 To pNumIters
        x = x + pSample(i, colIndex)
    Next i
    pSampleSum = x
End Function

Public Property Get SampleMean(colIndex As Long) As Double
    SampleMean = pSampleSum(colIndex) / pNumIters
End Property

Public Property Get CumulativeAbscissa(rowIndex As Long, colIndex As Long) As Double
    CumulativeAbscissa = pOrderedSample(rowIndex, colIndex)
End Property

Public Property Get CumulativeOrdinate(rowIndex As Long, colIndex As Long) As Double
    CumulativeOrdinate = rowIndex / (NumIters + 1)
End Property

Public Property Get Variance(colIndex As Long) As Double
    w = 0
    i = 0
    Do
        v = 0
        Do
            i = i + 1
            v = v + pSample(i, colIndex) ^ 2
        Loop Until (v >= 9E+300 Or i = pNumIters)
        w = w + (v / pNumIters)
    Loop Until i = pNumIters
    Variance = w - (pSampleSum(colIndex) / pNumIters) ^ 2
End Property



'===========================================================================================================================================
' Math Functions
'===========================================================================================================================================


Attribute VB_Name = "mc_maths_functions"
Option Explicit
Option Compare Text

' Excluding the function NumberOfArrayDimensions, Copyright 2015 Howard J Rudd
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, this software
' is distributed on an "AS IS" BASIS WITHOUT WARRANTIES OR CONDITIONS OF
' ANY KIND, either express or implied, not even for MERCHANTABILITY or
' FITNESS FOR A PARTICULAR PURPOSE. See the License for the specific language
' governing permissions and limitations under the License. You are free to use
' this code as you wish within the provisions of the license but it is your
' responsibility to test it and ensure it is fit for the use to which you
' intend to put it.
' _____________________________________________________________________________
'
' This module contains the following functions that perform mathematical
' calculations needed for Monte Carlo risk analysis:
'
'    1. arrayrank(vArray)
'    2. chol(A)
'    3. expon(x)
'    4. finvs(F, S)
'    5. ic(x, C)
'    6. matmult(A, B)
'    7. mattransmult(A, B
'    8. normalscores(n, m)
'    9. NumberOfArrayDimensions(Arr)
'   10. qs1dd(inLow, inHi, keyArray, otherArray)
'   11. qs1ds(inLow, inHi, keyArray)
'   12. qs2dd(inLow, inHi, keyArray, column, otherArray)
'   13. qs2ds(inLow, inHi, keyArray, column)
'   14. quicksort(keyArray, Optional column, Optional otherArray)
'   15. shuffle(vArray, Optional column)
'   16. unifun(n)
'
'______________________________________________________________________________

Public Function arrayrank(vArray() As Double) As Long()

' Usage: Y = arrayrank(vArray)
'
' Returns an array of longs representing the the rank orders of the elements
' of vArray. If vArray is two-dimensional, then it returns an array with
' the same number of rows and columns with the columns containing the ranks
' of the corresponding columns of "vArray".

Dim vDims As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long
Dim n As Long
Dim vRank() As Long
Dim tmpRank() As Long

' Check that dimensionality of vArray is either 1 or 2
vDims = NumberOfArrayDimensions(vArray)
If Not (vDims = 1 Or vDims = 2) Then
    Err.Raise Number:=vbObjectError + 7, _
              Source:="arrayrank", _
              Description:="Input argument not one or two dimensional"
    Debug.Print "Arrayrank input argument not one or two dimensional"
    Debug.Print vDims
    Exit Function
End If

' Check that array indices are numbered from 1
If vDims = 1 Then
    If Not (LBound(vArray) = 1) Then
        Err.Raise Number:=vbObjectError + 8, _
                  Source:="arrayrank", _
                  Description:="One dimensional array input argument not numbered from 1"
        Debug.Print "arrayrank input argument start index = " & LBound(vArray)
        Exit Function
    End If
    n = UBound(vArray)
ElseIf vDims = 2 Then
    If Not (LBound(vArray, 1) = 1 And LBound(vArray, 2) = 1) Then
        Err.Raise Number:=vbObjectError + 9, _
                  Source:="arrayrank", _
                  Description:="Two dimensional array input argument not numbered from 1"
        Debug.Print "column index starts from" & LBound(vArray, 1) & " and row index starts from" & LBound(vArray, 1)
    End If
    n = UBound(vArray, 1) - LBound(vArray, 1) + 1
    m = UBound(vArray, 2) - LBound(vArray, 2) + 1
End If

' Create array of consecutive longs from startIndex to endIndex
If vDims = 1 Then
    ReDim vRank(1 To n)
    ReDim tmpRank(1 To n)
ElseIf vDims = 2 Then
    ReDim vRank(1 To n, 1 To m)
    ReDim tmpRank(1 To n)
End If

' Pass this array to QuickSort along with the array to be ranked
If vDims = 1 Then
    For i = 1 To n
        tmpRank(i) = i
    Next i
        Call quicksort(keyArray:=vArray, otherArray:=tmpRank)
    For i = 1 To n
        vRank(tmpRank(i)) = i
    Next i
ElseIf vDims = 2 Then
    For j = 1 To m
        For i = 1 To n
            tmpRank(i) = i
        Next i
            Call quicksort(keyArray:=vArray, Column:=j, otherArray:=tmpRank)
        For i = 1 To n
            vRank(tmpRank(i), j) = i
        Next i
    Next j
End If

arrayrank = vRank

End Function

Public Function chol(A() As Double) As Double()

' Usage: Y = chol(A)
'
' Returns the upper triangular Cholesky root of A. A must by square, symmetric
' and positive definite.

Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long
Dim n As Long
Dim G() As Double

'Determine the number of rows, n, and number of columns, m, in A
n = UBound(A, 1) - LBound(A, 1) + 1
m = UBound(A, 2) - LBound(A, 2) + 1

'TESTS

' Check that A is indexed from 1
If Not (LBound(A, 1) = 1 And LBound(A, 2) = 1) Then
    Err.Raise Number:=vbObjectError + 10, _
        Source:="chol", _
        Description:="Matrix not indexed from 1"
    Debug.Print "column index starts from" & LBound(A, 1) & " and row index starts from" & LBound(A, 1)
End If

'Check that A is square
If Not (n = m) Then
    Debug.Print "Attempted to perform Cholesky factorisation on a matrix" _
                & "that is not square"
    Err.Raise Number:=vbObjectError + 11, _
              Source:="chol", _
              Description:="Matrix not square"
    Exit Function
End If

'Check that A is at least 2 x 2
If Not WorksheetFunction.Min(n, m) >= 2 Then
    Debug.Print "Attempted to perform Cholesky factorisation on a 1 x 1 matrix"
    Err.Raise Number:=vbObjectError + 12, _
              Source:="chol", _
              Description:="Input matrix only has one element"
    Exit Function
End If

'Check that A is symmetric
For i = 1 To n
    For j = 1 To m
        If Not A(i, j) = A(j, i) Then
               Debug.Print "Attempted to perform Cholesky factorisation on a matrix" _
                         & "that is not symmetric"
                Err.Raise Number:=vbObjectError + 13, _
                          Source:="chol", _
                          Description:="Matrix not symmetric"
            Exit Function
        End If
    Next j
Next i

' Check that the first diagonal element of A is >= 0
If A(1, 1) <= 0 Then
    Debug.Print "Attempted to perform Cholesky factorisation on a matrix" & _
                "that is not positive definite"
    Err.Raise Number:=vbObjectError + 14, _
                Source:="chol", _
                Description:="Matrix not positive definite"
    Exit Function
End If

' END OF TESTS (almost). Actual maths starts here!

ReDim G(1 To n, 1 To m)

' Calculate 1st element
G(1, 1) = Sqr(A(1, 1))

' Calculate remainder of 1st row
For j = 2 To m
    G(1, j) = A(1, j) / G(1, 1)
Next j

' Calculate remaining rows
For i = 2 To n
' Calculate diagonal element of row i
    G(i, i) = A(i, i)
    For k = 1 To i - 1
        G(i, i) = G(i, i) - G(k, i) * G(k, i)
    Next k
' Check that g(i,i) is > 0
    If G(i, i) <= 0 Then
        Debug.Print "Attempted to perform Cholesky factorisation on a matrix" & _
                    "that is not positive definite"
        Err.Raise Number:=vbObjectError + 15, _
                  Source:="chol", _
                  Description:="Matrix not positive definite"
        Exit Function
    End If
    G(i, i) = Sqr(G(i, i))
' Calculate remaining elements of row i
    For j = i + 1 To m
        G(i, j) = A(i, j)
        For k = 1 To i - 1
            G(i, j) = G(i, j) - G(k, i) * G(k, j)
        Next k
        G(i, j) = G(i, j) / G(i, i)
    Next j
Next i

' Calculate lower triangular half of G
For i = 2 To n
    For j = 1 To i - 1
        G(i, j) = 0
    Next j
Next i

chol = G

End Function

Public Function expon(x As Double) As Integer

' Usage: Y = expon(x)
'
' Returns the exponent to the base 10 of x. That is the value 'b' such that
' x = a * 10^b, where 1<= a < 10.

Dim y As Double

y = Log(x) / Log(10)
expon = Int(y)

End Function

Public Function finvs(F, S) As Double()

' Usage: Y = finvs(F, S)
'
' Returns the product of F^-1 and S for F and S both upper triangular. Actually
' solves FZ = S, i.e. finds Z such that FZ = S.

Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long
Dim n As Long

Dim nRowsF As Long
Dim nColsF As Long
Dim nRowsS As Long
Dim nColsS As Long

Dim z() As Double
Dim w As Double

nRowsF = UBound(F, 1) - LBound(F, 1) + 1
nColsF = UBound(F, 2) - LBound(F, 2) + 1
nRowsS = UBound(S, 1) - LBound(S, 1) + 1
nColsS = UBound(S, 2) - LBound(S, 2) + 1

' TESTS

' Check that F is indexed from 1
If Not (LBound(F, 1) = 1 And LBound(F, 2) = 1) Then
    Err.Raise Number:=vbObjectError + 16, _
        Source:="finvs", _
        Description:="Matrix F not indexed from 1"
    Debug.Print "column index starts from" & LBound(F, 1) & " and row index starts from" & LBound(F, 1)
End If

' Check that S is indexed from 1
If Not (LBound(S, 1) = 1 And LBound(S, 2) = 1) Then
    Err.Raise Number:=vbObjectError + 17, _
        Source:="finvs", _
        Description:="Matrix S not indexed from 1"
    Debug.Print "column index starts from" & LBound(S, 1) & " and row index starts from" & LBound(F, 1)
End If

' Test whether F is square
If Not nRowsF = nColsF Then
    Debug.Print "Matrix F is not square"
    Err.Raise Number:=vbObjectError + 18, _
              Source:="finvs", _
              Description:="Matrix F is not square"
    Exit Function
End If

' Test whether S is square
If Not nRowsS = nColsS Then
    Debug.Print "Matrix S is not square"
    Err.Raise Number:=vbObjectError + 19, _
              Source:="finvs", _
              Description:="Matrix S is not square"
    Exit Function
End If

' Test whether F and S have same dimensions
If Not nRowsF = nRowsS And nColsF = nColsS Then
    Debug.Print "Matrices F and S have different dimensions"
    Err.Raise Number:=vbObjectError + 20, _
              Source:="finvs", _
              Description:="Matrices F and S have different dimensions"
    Exit Function
End If

' Test whether F is upper triangular
For i = 1 To nRowsF
    For j = 1 To i - 1
        If Not (F(i, j) = 0 And (Not F(j, i) = 0)) Then
            Debug.Print "Matrix F is not upper triangular"
            Err.Raise Number:=vbObjectError + 21, _
                      Source:="finvs", _
                      Description:="Matrix F not upper triangular"
            Exit Function
        End If
    Next j
Next i

' Test whether S is upper triangular
For i = 2 To nRowsS
    For j = 1 To i - 1
        If S(i, j) > 10 ^ (-16) Then
            Debug.Print "Matrix S is not upper triangular"
            Err.Raise Number:=vbObjectError + 22, _
                      Source:="finvs", _
                      Description:="Matrix S not upper triangular"
            Exit Function
        End If
    Next j
Next i

' Test whether F has all non-zero diagonal elements
For i = 1 To nRowsF
        If F(i, i) = 0 Then
            Debug.Print "Matrix F has at least one zero diagonal element and so is not invertible"
            Err.Raise Number:=vbObjectError + 23, _
                      Source:="finvs", _
                      Description:="Matrix F has at least one zero diagonal element and so is not invertible"
            Exit Function
        End If
Next i

' END OF TESTS. Actual maths starts here!

n = nRowsF

ReDim z(1 To n, 1 To n)

' Construct the nth row of Z
For j = 1 To n - 1
    z(n, j) = 0
Next j
    z(n, n) = S(n, n) / F(n, n)

' Construct the rows of Z above the nth
For i = n - 1 To 1 Step -1
    For j = 1 To n
        w = 0
        For k = i + 1 To n
            w = w + F(i, k) * z(k, j)
        Next k
        z(i, j) = (S(i, j) - w) / F(i, i)
    Next j
Next i

finvs = z

End Function

Public Function ic(Xascending() As Double, C() As Double) As Double()

' Usage: y = ic(Xascending, C)
'
' Performs the Iman-Conover method on Xind and returns a matrix Xcorr with the same
' dimensions as Xind.
'
' Xascending is an n-instance sample from an m-element random row-vector, with
' each column sorted in ascending order.
'
' If the columns of Xascending are in ascending order then the correlation
' matrix of Xcorr will be approximately equal to C. This function does not test
' Xascending to check that its columns are in ascending order. To do so would
' incur too large a computational burden.
'
' C must be square, symmetric and positive definite.
'
' The number of rows and columns of C must equal the number of columns Xascending.

Dim nRowsC As Long
Dim nColsC As Long
Dim nRowsX As Long
Dim nColsX As Long

nRowsC = UBound(C, 1) - LBound(C, 1) + 1
nColsC = UBound(C, 2) - LBound(C, 2) + 1
nRowsX = UBound(Xascending, 1) - LBound(Xascending, 1) + 1
nColsX = UBound(Xascending, 2) - LBound(Xascending, 2) + 1

Dim i As Long
Dim j As Long
Dim k As Long

Dim EX() As Double
ReDim EX(1 To nColsX, 1 To nColsX)
Dim FX() As Double
ReDim FX(1 To nColsX, 1 To nColsX)
Dim ZX() As Double
ReDim ZX(1 To nColsX, 1 To nColsX)
Dim S() As Double
ReDim S(1 To nColsX, 1 To nColsX)

Dim MX() As Double
ReDim MX(1 To nRowsX, 1 To nColsX)
Dim TX() As Double
ReDim TX(1 To nRowsX, 1 To nColsX)
Dim YX() As Double
ReDim YX(1 To nRowsX, 1 To nColsX)
Dim ranks() As Long
ReDim ranks(1 To nRowsX, 1 To nColsX)

' TESTS!

' Test if C is square
If Not nRowsC = nColsC Then
    MsgBox Title:="Iman-Conover Function", _
           prompt:="Correlation matrix is not square"
End If
  
' Test if C is symmetric
For i = 1 To nRowsC
    For j = i To nColsC
        If Abs(C(i, j) - C(j, i)) >= 10 ^ (-16) Then
            MsgBox Title:="Iman-Conover Function", _
                prompt:="Correlation matrix is not symmetric"
            Exit Function
        End If
    Next j
Next i

' Test if the number of rows of C is greater than the number of columns of X
If nRowsC > nColsX Then
    MsgBox Title:="Iman-Conover Function", _
           prompt:="Correlation matrix too large"
End If

' Test if the number of rows of C is less than the number of columns of X
If nRowsC < nColsX Then
    MsgBox Title:="Iman-Conover Function", _
           prompt:="Correlation matrix too small"
End If

' END OF TESTS: Actual maths starts here!

' Calculate the upper triangular Cholesky root of C
S = chol(C)

' Calculate the matrix, MX, of "normal scores"
MX = normalscores(nRowsX, nColsX)

' Calculate the matrix EX = MX' * MX
EX = mattransmult(MX, MX)

' Calculate Fx, the Cholesky root of Ex.
FX = chol(EX)

' Calculate ZX = FX^{-1) * S
ZX = finvs(FX, S)

' Calculate TX, the reordered matrix of "scores".
TX = matmult(MX, ZX)

' Calculate the rank orders of TX.
ranks = arrayrank(TX)

' Reorder columns of X to match T.
For j = 1 To nColsX
    For k = 1 To nRowsX
        YX(k, j) = Xascending(ranks(k, j), j)
    Next k
Next j

ic = YX

End Function

Public Function matmult(A, B) As Double()

'Usage: C = matmult(A, B)
'
' Returns the product of A and B. Start indices of A and B can be arbitrary but
' start indices of the product C are both 1.

Dim startIndexRowsA As Long
Dim endIndexRowsA As Long
Dim startIndexColsA As Long
Dim endIndexColsA As Long
Dim startIndexRowsB As Long
Dim endIndexRowsB As Long
Dim startIndexColsB As Long
Dim endIndexColsB As Long
Dim nRowsA As Long
Dim nColsA As Long
Dim nRowsB As Long
Dim nColsB As Long
Dim nRowsC As Long
Dim nColsC As Long
Dim i As Long
Dim j As Long
Dim k As Long

startIndexRowsA = LBound(A, 1)
endIndexRowsA = UBound(A, 1)
startIndexColsA = LBound(A, 2)
endIndexColsA = UBound(A, 2)
startIndexRowsB = LBound(B, 1)
endIndexRowsB = UBound(B, 1)
startIndexColsB = LBound(B, 2)
endIndexColsB = UBound(B, 2)

nRowsA = endIndexRowsA - startIndexRowsA + 1
nColsA = endIndexColsA - startIndexColsA + 1
nRowsB = endIndexRowsB - startIndexRowsB + 1
nColsB = endIndexColsB - startIndexColsB + 1

' Test that the two matrices are conformable
If Not nColsA = nRowsB Then
    Debug.Print "Attempted to multiply non conformable matrices"
    Err.Raise Number:=vbObjectError + 24, _
              Source:="matmult", _
              Description:="Matrices not conformable"
    Exit Function
End If
          
nRowsC = nRowsA
nColsC = nColsB

Dim C() As Double
ReDim C(1 To nRowsC, 1 To nColsC)

For i = 1 To nRowsC
    For j = 1 To nColsC
        C(i, j) = 0
        For k = 1 To nColsA
            C(i, j) = C(i, j) + A(i + startIndexRowsA - 1, k + startIndexColsA - 1) * _
                                B(k + startIndexRowsB - 1, j + startIndexColsB - 1)
        Next k
    Next j
Next i

matmult = C

End Function

Public Function mattransmult(A, B) As Double()

'Usage: C = mattransmult(A, B)
'
' Returns the product of A-transpose and B. Start indices of A and B
' can be arbitrary but start indices of the product C are both 1.

Dim startIndexRowsA As Long
Dim endIndexRowsA As Long
Dim startIndexColsA As Long
Dim endIndexColsA As Long
Dim startIndexRowsB As Long
Dim endIndexRowsB As Long
Dim startIndexColsB As Long
Dim endIndexColsB As Long
Dim nRowsA As Long
Dim nColsA As Long
Dim nRowsB As Long
Dim nColsB As Long
Dim nRowsC As Long
Dim nColsC As Long
Dim i As Long
Dim j As Long
Dim k As Long

startIndexRowsA = LBound(A, 1)
endIndexRowsA = UBound(A, 1)
startIndexColsA = LBound(A, 2)
endIndexColsA = UBound(A, 2)
startIndexRowsB = LBound(B, 1)
endIndexRowsB = UBound(B, 1)
startIndexColsB = LBound(B, 2)
endIndexColsB = UBound(B, 2)

nRowsA = endIndexRowsA - startIndexRowsA + 1
nColsA = endIndexColsA - startIndexColsA + 1
nRowsB = endIndexRowsB - startIndexRowsB + 1
nColsB = endIndexColsB - startIndexColsB + 1

' Test that the two matrices are conformable
If Not nRowsA = nRowsB Then
    Debug.Print "Attempted to multiply non conformable matrices"
    Err.Raise Number:=vbObjectError + 25, _
              Source:="matmult", _
              Description:="Matrices not conformable"
    Exit Function
End If
          
nRowsC = nColsA
nColsC = nColsB

Dim C() As Double
ReDim C(1 To nRowsC, 1 To nColsC)

For i = 1 To nColsA
    For j = 1 To nColsB
        C(i, j) = 0
        For k = 1 To nRowsA
            C(i, j) = C(i, j) + A(k + startIndexRowsA - 1, i + startIndexColsA - 1) * _
                                B(k + startIndexRowsB - 1, j + startIndexColsB - 1)
        Next k
    Next j
Next i

mattransmult = C

End Function

Public Function normalscores(n, m)

Dim Omega() As Double

ReDim Omega(1 To n, 1 To m)

Dim i As Long
Dim j As Long
Dim k As Long
Dim x As Double
Dim NormalisationFactor As Double

x = 0
For i = 1 To n \ 2
    Omega(i, 1) = WorksheetFunction.NormInv((i / (n + 1)), 0, 1)
    x = x + Omega(i, 1) ^ 2
Next i

NormalisationFactor = Sqr(2 * x / n)

For i = 1 To n \ 2
    Omega(i, 1) = Omega(i, 1) / NormalisationFactor
Next i

k = Int(n / 2) + 1

If Not 2 * Int(n / 2) = n Then
    Omega(k, 1) = 0
    For i = k + 1 To n
        Omega(i, 1) = -Omega(n - i + 1, 1)
    Next i
Else
    For i = k To n
        Omega(i, 1) = -Omega(n - i + 1, 1)
    Next i
End If

For j = 2 To m
    For i = 1 To n
        Omega(i, j) = Omega(i, 1)
    Next i
Next j

For i = 1 To m
    Call shuffle(Omega, i)
Next i

normalscores = Omega

ReDim Omega(0)

End Function

Public Function NumberOfArrayDimensions(Arr As Variant) As Long
' This function is from http://www.cpearson.com/Excel/VBAArrays.htm

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim Res As Long
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    Res = UBound(Arr, Ndx)
Loop Until Err.Number <> 0

NumberOfArrayDimensions = Ndx - 1

End Function

Private Function qs1ds(inLow, inHi, keyArray)

' Usage: quicksort(inLow, inHi, keyArray)
'
' "qs1ds" = quick sort one dimensional single array
'
' Sorts keyArray in place between indices inLow and inHi, keyArray must be of
' type Double
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

Dim tmpLow As Long
Dim tmpHi As Long
Dim pivot As Double
Dim keyTmpSwap As Double

tmpLow = inLow
tmpHi = inHi
pivot = keyArray((inLow + inHi) \ 2)

Do While (tmpLow <= tmpHi)
    Do While keyArray(tmpLow) < pivot And tmpLow < inHi
        tmpLow = tmpLow + 1
    Loop
    Do While keyArray(tmpHi) > pivot And tmpHi > inLow
        tmpHi = tmpHi - 1
    Loop
    If (tmpLow <= tmpHi) Then
        keyTmpSwap = keyArray(tmpLow)
        keyArray(tmpLow) = keyArray(tmpHi)
        keyArray(tmpHi) = keyTmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Loop

If (inLow < tmpHi) Then qs1ds inLow, tmpHi, keyArray
If (tmpLow < inHi) Then qs1ds tmpLow, inHi, keyArray

End Function

Private Function qs1dd(inLow, inHi, keyArray, otherArray)

' Usage: qs1dd(inLow, inHi, keyArray, otherArray)
'
' "qs1dd" = quick sort one dimensional double array
'
' Sorts keyArray in place between indices inLow and inHi.
'
' An array, "otherArray" is sorted in parallel, also in-place. "otherArray" must
' be one dimensional and have the same start and end indices as the first dimen-
' sion of "keyArray".
'
' "keyArray" must be of type Double, "otherArray" must be of type long.
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

Dim pivot As Double
Dim tmpLow As Long
Dim tmpHi As Long
Dim keyTmpSwap As Double
Dim otherTmpSwap As Long
Dim keyDims As Long
Dim otherDims As Long

tmpLow = inLow
tmpHi = inHi
pivot = keyArray((inLow + inHi) \ 2)

Do While (tmpLow <= tmpHi)

    Do While keyArray(tmpLow) < pivot And tmpLow < inHi
        tmpLow = tmpLow + 1
    Loop
    
    Do While keyArray(tmpHi) > pivot And tmpHi > inLow
        tmpHi = tmpHi - 1
    Loop
    
    If (tmpLow <= tmpHi) Then
    
        keyTmpSwap = keyArray(tmpLow)
        otherTmpSwap = otherArray(tmpLow)
        
        keyArray(tmpLow) = keyArray(tmpHi)
        otherArray(tmpLow) = otherArray(tmpHi)
        
        keyArray(tmpHi) = keyTmpSwap
        otherArray(tmpHi) = otherTmpSwap
        
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
        
    End If
    
Loop

If (inLow < tmpHi) Then qs1dd inLow, tmpHi, keyArray, otherArray
If (tmpLow < inHi) Then qs1dd tmpLow, inHi, keyArray, otherArray

End Function

Private Function qs2ds(inLow As Long, inHi As Long, keyArray() As Double, Column As Long)

' Usage: qs2ds(inLow, inHi, keyArray, column)
'
' "qs2ds" = quick sort two dimensional single array
'
' Sorts a column of "keyArray", in place ' between indices "inLow" and "inHi".
' "keyArray" must be is two-dimensional (i.e. a matrix). Only the column
' specified by the optional argument "column" will be sorted, the other columns
' are left untouched.
'
' "keyArray" must be of type Double, column and "column" must be of type
' long
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

Dim pivot As Double
Dim tmpLow As Long
Dim tmpHi As Long
Dim keyTmpSwap As Double
Dim keyDims As Long

tmpLow = inLow
tmpHi = inHi
pivot = keyArray((inLow + inHi) \ 2, Column)

Do While (tmpLow <= tmpHi)
    Do While keyArray(tmpLow, Column) < pivot And tmpLow < inHi
        tmpLow = tmpLow + 1
    Loop
    Do While keyArray(tmpHi, Column) > pivot And tmpHi > inLow
        tmpHi = tmpHi - 1
    Loop
    If (tmpLow <= tmpHi) Then
        keyTmpSwap = keyArray(tmpLow, Column)
        keyArray(tmpLow, Column) = keyArray(tmpHi, Column)
        keyArray(tmpHi, Column) = keyTmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Loop

If (inLow < tmpHi) Then qs2ds inLow, tmpHi, keyArray, Column
If (tmpLow < inHi) Then qs2ds tmpLow, inHi, keyArray, Column

End Function

Private Function qs2dd(inLow, inHi, keyArray, Column, otherArray)

' Usage: qs2dd(inLow, inHi, keyArray, column, otherArray)
'
' "qs1dd" = quick sort two dimensional double array
'
' Sorts "keyArray" and "otherArray" in place between indices inLow and inHi.
'
' An array, "otherArray" is sorted in parallel, also in-place. "otherArray" must
' be one dimensional and have the same start and end indices as the first dimen-
' sion of "keyArray".
'
' "keyArray" must be of type Double, "column" and "otherArray" must be of type
' long.
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

Dim pivot As Double
Dim tmpLow As Long
Dim tmpHi As Long
Dim keyTmpSwap As Double
Dim otherTmpSwap As Long
Dim keyDims As Long
Dim otherDims As Long

tmpLow = inLow
tmpHi = inHi
pivot = keyArray((inLow + inHi) \ 2, Column)

Do While (tmpLow <= tmpHi)

    Do While keyArray(tmpLow, Column) < pivot And tmpLow < inHi
        tmpLow = tmpLow + 1
    Loop
    
    Do While keyArray(tmpHi, Column) > pivot And tmpHi > inLow
        tmpHi = tmpHi - 1
    Loop
    
    If (tmpLow <= tmpHi) Then
    
        keyTmpSwap = keyArray(tmpLow, Column)
        otherTmpSwap = otherArray(tmpLow)
        
        keyArray(tmpLow, Column) = keyArray(tmpHi, Column)
        otherArray(tmpLow) = otherArray(tmpHi)
        
        keyArray(tmpHi, Column) = keyTmpSwap
        otherArray(tmpHi) = otherTmpSwap
        
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
        
    End If
    
Loop

If (inLow < tmpHi) Then qs2dd inLow, tmpHi, keyArray, Column, otherArray
If (tmpLow < inHi) Then qs2dd tmpLow, inHi, keyArray, Column, otherArray

End Function

Public Function quicksort(keyArray() As Double, Optional Column As Long, Optional otherArray)

' Usage: quicksort(keyArray, column, otherArray)
'
' Sorts keyArray in place. If keyArray is two-dimensional (i.e. a matrix) then
' only the column specified by the optional argument "column" will be sorted.
'
' An optional "otherArray" can be sorted in parallel, also in-place. If
' "otherArray" is used it must be one-dimensional and have the same start and
' end indices as the columns of "keyArray".
'
' "keyArray" must be of type Double, "column" and otherArray must be of type
' long

Dim keyDims As Long
Dim otherDims As Long
Dim inLow As Long
Dim inHi As Long
Dim i As Long

' TESTS

' Check that the dimensionality of "keyArray" is either 1 or 2
keyDims = NumberOfArrayDimensions(keyArray)
If Not (keyDims = 1 Or keyDims = 2) Then
    Debug.Print "input argument not one or two dimensional"
    Exit Function
End If

If Not IsMissing(otherArray) Then
' Check that "otherArray" is one-dimensional
    otherDims = NumberOfArrayDimensions(otherArray)
    If Not otherDims = 1 Then
        Debug.Print "'otherArray' not one-dimensional"
        Exit Function
    End If
    If keyDims = 1 Then
' Check that "keyArray" and "otherArray" are conformable
        If Not ( _
                    UBound(keyArray) = UBound(otherArray) And _
                    LBound(keyArray) = LBound(otherArray) _
                ) Then
                Debug.Print "'keyArray' and 'otherArray' not conformable"
                Exit Function
        End If
    ElseIf keyDims = 2 Then
' Check that the argument "column" has been supplied
        If IsMissing(Column) Then
            Debug.Print "'column' argument not passed"
            Exit Function
        End If
' Check that "otherArray" is conformable to columns of "keyArray"
        If Not ( _
                UBound(keyArray, 1) = UBound(otherArray) And _
                LBound(keyArray, 1) = LBound(otherArray) _
            ) Then
            Debug.Print "'keyArray' and 'otherArray' not conformable"
            Exit Function
        End If
    End If
End If

' Check that the argument "column" points to one of the columns of "keyArray"
If Not (LBound(keyArray, 2) <= Column And Column <= UBound(keyArray, 2)) Then
    ' ERROR: Argument "column" does not point to one of the columns of "keyArray"
    Exit Function
End If

' END OF TESTS

' Calculate "inHi" and "inLow"

If keyDims = 1 Then
    inLow = LBound(keyArray)
    inHi = UBound(keyArray)
ElseIf keyDims = 2 Then
    inLow = LBound(keyArray, 1)
    inHi = UBound(keyArray, 1)
End If

' Call appropriate sort function

If keyDims = 1 And IsMissing(otherArray) Then
    Call qs1ds(inLow, inHi, keyArray)
        
ElseIf keyDims = 1 And Not IsMissing(otherArray) Then
    Call qs1dd(inLow, inHi, keyArray, otherArray)

ElseIf keyDims = 2 And IsMissing(otherArray) Then
    Call qs2ds(inLow, inHi, keyArray, Column)

ElseIf keyDims = 2 And Not IsMissing(otherArray) Then
    Call qs2dd(inLow, inHi, keyArray, Column, otherArray)

End If

End Function

Public Function shuffle(vArray, Optional Column)

' Usage: shuffle(vArray, column)
'
' shuffles vArray in place randomly. If vArray is 2-dimensional, then it shuffles
' the column specified by the optional variable "column"

Dim j As Long
Dim k As Long
Dim startIndex As Long
Dim endIndex As Long
Dim vDims As Long

Dim temp As Double

' Check that dimensionality of "vArray" is either 1 or 2
vDims = NumberOfArrayDimensions(vArray)
If Not (vDims = 1 Or vDims = 2) Then
    Debug.Print "input argument not one or two dimensional"
    Debug.Print vDims
    Exit Function
End If

If vDims = 1 Then
    startIndex = LBound(vArray)
    endIndex = UBound(vArray)
    j = endIndex
    Do While j > startIndex
        k = CLng(startIndex + Rnd() * (j - startIndex))
        temp = vArray(j)
        vArray(j) = vArray(k)
        vArray(k) = temp
        j = j - 1
    Loop
ElseIf vDims = 2 Then
    If IsMissing(Column) Then
        'ERROR: Argument 'column' not supplied
        Exit Function
    End If
' Check that the argument "column" points to one of the columns of "vArray"
    If Not (LBound(vArray, 2) <= Column And Column <= UBound(vArray, 2)) Then
        ' ERROR: Argument "column" does not point to one of the columns of "vArray"
        Exit Function
    End If
    startIndex = LBound(vArray, 1)
    endIndex = UBound(vArray, 1)
    j = endIndex
    Do While j > startIndex
        k = CLng(startIndex + Rnd() * (j - startIndex))
        temp = vArray(j, Column)
        vArray(j, Column) = vArray(k, Column)
        vArray(k, Column) = temp
        j = j - 1
    Loop
End If

End Function

Public Function unifun(ByVal n As Long) As Double()

' Usage: y = unifun(n)
'
' Returns an array of n random numbers in ascending order, the ith member of
' which is a uniformly distributed on the range (i - 1)/n <= x < i/n, except
' for the 1st member, which is uniformly distributed on the range
' 0 < x < 1/n.

Dim U As Double
Dim y() As Double
ReDim y(1 To n)
Dim i As Long

' First member of the sequence
Do
    U = Rnd()
Loop Until U <> 0
y(1) = U / n
    
' Remaining members of the sequence
For i = 2 To n
    y(i) = (Rnd() + i - 1) / n
Next i
    
unifun = y

End Function


'=====================================================================================================================================
' Quantile Functions
'======================================================================================================================================


Option Explicit
Option Compare Text

' Copyright 2015 Howard J Rudd
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, this software
' is distributed on an "AS IS" BASIS WITHOUT WARRANTIES OR CONDITIONS OF
' ANY KIND, either express or implied, not even for MERCHANTABILITY or
' FITNESS FOR A PARTICULAR PURPOSE. See the License for the specific language
' governing permissions and limitations under the License. You are free to use
' this code as you wish within the provisions of the license but it is your
' responsibility to test it and ensure it is fit for the use to which you
' intend to put it.
'
' _____________________________________________________________________________


' This module contains functions that return the quantiles of the random
' variables that are needed in the simulations. These are in a separate module
' to make it easier for the user to add new ones as required. The following
' functions are included:
'
'   01. TriangularInv(y, left, middle, right)
'   02. UniformInv(y, left, right)
'
'______________________________________________________________________________


Public Function TriangularInv(y, left, middle, right) As Double

' Returns the quantile function of a random variable distributed according to
' the triangular distribution with corners at x = a, x = b and x = c.

Dim tmp As Variant
Dim cond12 As Boolean
Dim cond23 As Boolean
Dim test1 As Boolean
Dim test2 As Boolean
Dim test3 As Boolean
Dim i As Integer

cond12 = left > middle
cond23 = middle > right

If cond12 Then
    MsgBox Title:="TriangularInv", _
           prompt:="Left corner to the right of top corner, " & vbCr & _
                   "corners reversed and calculation continued."
ElseIf cond23 Then
    MsgBox Title:="TriangularInv", _
           prompt:="Top corner to the right of right corner, " & vbCr & _
                   "corners reversed and calculation continued."
End If

' The following loop sorts a, b and c into ascending order. This is done because
' it is assumed that the user intended to enter the same numbers in ascending
' order but made a mistake. The error is flagged via a message box, but the
' calculation continues. It is, however, possible that the user made a different
' mistake, such as entering one or more of the values incorrectly.

Do While cond12 Or cond23
    If cond12 Then
        tmp = left
        left = middle
        middle = tmp
    ElseIf cond23 Then
        tmp = middle
        middle = right
        right = tmp
    End If
    cond12 = left > middle
    cond23 = middle > right
Loop

' The variable isArrayy stores True if y is an array and False if not. Avoids
' the need to evaluate IsArray(y) multiple times.

test1 = left = middle
test2 = middle = right
test3 = left = right

If test1 And (Not test2) Then
' Situation 1: The triangle is right angled with the right angle on the left.
    TriangularInv = right - (right - left) * Sqr(1 - y)

ElseIf test2 And (Not test1) Then
' Situation 2: b and c are equal. The triangle is right angled with the right
' angle on the right.
    TriangularInv = left + (middle - left) * Sqr(y)

ElseIf test1 And test2 And test3 Then
' Situation 3: the x-coordinates of all the triangle's corners coincide. This is
' an error situation. The user has entered them incorrectly. Raise an error
' message, but still generate output. The output will be an array of constants
' equal to a = b = c. No randomness involved.
    TriangularInv = left

ElseIf (Not test1) And (Not test2) And (Not test3) Then
' Situation 4: the happy case.
    If y <= (middle - left) / (right - left) Then
        TriangularInv = left + Sqr(y * (right - left) * (middle - left))
    Else
        TriangularInv = right - Sqr((1 - y) * (right - left) * (right - middle))
    End If
Else
' Can't imagine what would be left if none of the above were true, just throw an
' exception if the programme ends up here.
        MsgBox Title:="TriangularInv", _
               prompt:="Er, something's wrong," & vbCr & _
                       "suggest check code."
End If

End Function

Public Function UniformInv(y, left, right) As Double

' Returns the quantile function of a random variable uniformly distributed on
' the range (a, b).
'
' Has to be decalred as Variant otherwise the argument y must either be always
' scalar or always an array, not sometimes one or sometimes the other. However,
' the function actually returns a double.

Dim tmp As Variant

' The following block tests to see whether b > a. If not, it genrates a warning
' then reverses the order of a and b and continues with the calculation.

If left > right Then
    MsgBox Title:="UniformInv", _
           prompt:="Lower end of input range greater than upper end, " & vbCr & _
                   "parameters reversed and calculation continued"
    tmp = left
    left = right
    right = tmp
End If

' The following block tests whether a = b and if so generates a warning

If left = right Then
    MsgBox Title:="UniformInv", _
           prompt:="Lower and upper ends of input range coincide." & vbCr & _
                   "Calculation continued but will return a constant"
End If

UniformInv = (right - left) * y + left

End Function

