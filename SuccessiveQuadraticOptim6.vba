' Math for Mere Mortals
' Excel Inline Optimization / Successive Quadratic Optimization v6
' http://mathformeremortals.wordpress.com

'Copyright (c) 2014, Math for Mere Mortals
'All rights reserved.
'Redistribution and use in source and binary forms, with or without modification, are permitted for noncommercial use provided that the following conditions are met:
' +  Redistributions of source code must retain the above copyright notice, this list of conditions, and the following disclaimer.
' +  Redistributions in binary form must reproduce the above copyright notice, this list of conditions, and the following disclaimer in the documentation and/or other materials provided with the distribution.

'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


' Portions of this code are adapted from Colin Legg's blog: http://colinlegg.wordpress.com/2014/01/14/vba-determine-all-precedent-cells-a-nice-example-of-recursion/


' THE 8 LINES BELOW NEED TO BE UNCOMMENTED AND COPIED TO THE WORKBOOK MODULE TO PREVENT EXCEL FROM CRASHING WHEN SAVING THE FILE.
'Private Sub Workbook_AfterSave(ByVal Success As Boolean)
'  ' Allow optimization calculations.
'  Call AllowTimerCall
'End Sub
'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'  ' Prevent optimization calculations during a save attempt.  This is necessary to prevent crashes when saving the workbook.
'  Call PreventTimerCall
'End Sub
' THE 8 LINES ABOVE NEED TO BE UNCOMMENTED AND COPIED TO THE WORKBOOK MODULE TO PREVENT EXCEL FROM CRASHING WHEN SAVING THE FILE.


Option Explicit
Option Base 1


Dim FunctionEvaluationCount As Long
Dim PreventTimer As Boolean
Const ShowDebugText = False
' There are two calculation modes.  One will be faster than the other, but it depends on your spreadsheet.  See the Math for Mere Mortals blog for details.
Const CalculationMode = "ByApplication" 'Valid values are: "ByApplication" (Default) or "ByPrecedents" or (fastest but risking incorrect results) "ByNaivePrecedents"


' RangeQueue is important because it contains all instances of the Minimize (or MinimizeNonnegative) function calls in the workbook.
' RangeQueue is a 2-dimensional array.
' The first dimension specifies the data type.
' 1: Range containing the Minimize or MinimizeNonnegative formula
' 2: Merit function value after the most recently completed optimization
' 3: Initial value used for the most recent completed optimization
' The second dimension indexes the Minimize function references.
' Despite its name, this data structure is not really implemented as a queue; it is unordered and cannot be empty.
Dim RangeQueue() As Variant
Dim TimerID As Long


' Windows API calls for SetTimer and KillTimer.
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms644906%28v=vs.85%29.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms644903%28v=vs.85%29.aspx
' Yes.  We have to convert to a Boolean type explicitly or VBA freaks out and doesn't evaluate the expression correctly.
#If CBool(Win64) Then
  ' Handles and pointers must have 64 bits so that Windows recognizes them.
  Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal HWnd As LongLong, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongLong) As Long
  Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal HWnd As LongLong, ByVal nIDEvent As Long) As Long
#Else
  ' Assume we are in 32-bit mode.  Handles and pointers are limited to 32 bits.
  Public Declare Function SetTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
  Public Declare Function KillTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long) As Long
#End If


' Prevent a recalculation from triggering the timer.
' This subprogram gives the Workbook module access to the PreventTimer variable.
Public Sub PreventTimerCall()
  PreventTimer = True
End Sub


' Allow a recalculation to trigger the timer.
' This subprogram gives the Workbook module access to the PreventTimer variable.
Public Sub AllowTimerCall()
  PreventTimer = False
End Sub


' Minimize TargetCell by changing ByChangingCell using initial guess InitialGuess.
' This function only starts a sequence of timers and triggers that result in minimization.
Public Function Minimize(Optional TargetCell As Range, Optional ByChangingCell As Range, Optional InitialGuess As Variant) As String
  Dim CallingRange As Range
  Dim SyntaxError As String
  ' Lazy initialization in case PreventTimer isn't initialized.
  If CStr(PreventTimer) = vbNullString Then
    PreventTimer = False
  End If
  ' Which cell recalculation triggered this?
  Set CallingRange = Application.Caller
  ' Check the syntax for this cell.
  SyntaxError = CheckSyntax(CallingRange)
  ' Is this formula valid?
  If SyntaxError = "" Then
    ' This module keeps a list of all cells with Minimize functions so that it can check them when automatic calculation is disabled.
    Call AddRange(CallingRange)
    ' Is this not already scheduled?  We don't want to schedule multiple timers for the same round of calculations.
    If TimerID = 0 Then
      ' Make sure the timer is allowed to be set.
      If PreventTimer Then
        If ShowDebugText Then
          Debug.Print "The system timer is blocked for an ongoing or pending calculation or a save operation."
        End If
      Else
        ' We cannot allow the user to trigger any events or calculations during the forthcoming cycle of optimizations.
        Application.Interactive = False
        ' It must not have been scheduled, so schedule it for 1 ms into the future.
        TimerID = SetTimer(0&, 0&, 1&, AddressOf ScheduleOptimizationServiceProc)
        ' Write a debug message.  Note that the cell referenced in the message is the one that triggered the optimization, not
        ' the first one to be optimized.
        If ShowDebugText Then
          Debug.Print "System Timer " & TimerID & " was instantiated for Optimize cell " & Application.Caller.Address & "."
        End If
      End If
    Else
      ' This is just a duplicate, a gift from Excel that, much like Aunt Edna's fruitcake, can be thrown away.
      If ShowDebugText Then
        Debug.Print "  No system timer for '" & Application.Caller.Worksheet.Name & "'!" & Application.Caller.Address & ".  This event will be ignored."
      End If
    End If
    ' Tell the user what we're doing...
    Minimize = "<<Minimize " & FormatRangeAddress(TargetCell, CallingRange) & " by changing " & FormatRangeAddress(ByChangingCell, CallingRange) & ">>"
  Else
    Minimize = SyntaxError
  End If
End Function


' Minimize TargetCell by changing ByChangingCell using initial guess InitialGuess.  ByChangingCell cannot be negative.
' This function only starts a sequence of timers and triggers that result in minimization.
Public Function MinimizeNonnegative(Optional TargetCell As Range, Optional ByChangingCell As Range, Optional InitialGuess As Variant) As String
  Dim CallingRange As Range
  Dim SyntaxError As String
  ' Lazy initialization in case PreventTimer isn't initialized.
  If CStr(PreventTimer) = vbNullString Then
    PreventTimer = False
  End If
  ' Which cell recalculation triggered this?
  Set CallingRange = Application.Caller
  ' Check the syntax for this cell.
  SyntaxError = CheckSyntax(CallingRange)
  ' Is this formula valid?
  If SyntaxError = "" Then
    ' This module keeps a list of all cells with Minimize functions so that it can check them when automatic calculation is disabled.
    Call AddRange(CallingRange)
    ' Is this not already scheduled?  We don't want to schedule multiple timers for the same round of calculations.
    If TimerID = 0 Then
      ' Make sure the timer is allowed to be set.
      If PreventTimer Then
        If ShowDebugText Then
          Debug.Print "The system timer is blocked for an ongoing or pending calculation or a save operation."
        End If
      Else
        ' We cannot allow the user to trigger any events or calculations during the forthcoming cycle of optimizations.
        Application.Interactive = False
        ' It must not have been scheduled, so schedule it for 1 ms into the future.
        TimerID = SetTimer(0&, 0&, 1&, AddressOf ScheduleOptimizationServiceProc)
        ' Write a debug message.  Note that the cell referenced in the message is the one that triggered the optimization, not
        ' the first one to be optimized.
        If ShowDebugText Then
          Debug.Print "System Timer " & TimerID & " was instantiated for Optimize cell " & Application.Caller.Address & "."
        End If
      End If
    Else
      ' This is just a duplicate, a gift from Excel that, much like Aunt Edna's fruitcake, can be thrown away.
      If ShowDebugText Then
        Debug.Print "  No system timer for '" & Application.Caller.Worksheet.Name & "'!" & Application.Caller.Address & ".  This event will be ignored."
      End If
    End If
    ' Tell the user what we're doing...
    MinimizeNonnegative = "<<Minimize " & FormatRangeAddress(TargetCell, CallingRange) & " by changing " & FormatRangeAddress(ByChangingCell, CallingRange) & " (nonnegative)>>"
  Else
    MinimizeNonnegative = SyntaxError
  End If
End Function


' Checks syntax and returns an error message if it finds an error (or an empty string if everything is ok)
' The function must be in the form Minimize(A1, B1, C) or MinimizeNonzero(A1, B1, C)
' A1 and B1 are each a reference to a single cell with nothing else.
' C is a single numerical value that Excel can evaluate.  C can include cell references but cannot contain any commas.
Function CheckSyntax(OptimCell As Range) As String
  Dim FunctionName As String
  Dim ArgumentData As String
  Dim FormulaParameters() As String
  Dim TargetCell As Variant
  Dim ByChangingCell As Variant
  Dim ErrorMessage As String
  Dim i As Long
  Dim ParameterCount As Long
  ' Is this Minimize or MinimizeNonzero?
  FunctionName = Mid(OptimCell.Formula, 2, InStr(OptimCell.Formula, "(") - 2)
  ' Figure out the three function arguments by parsing the text in the cell formula...
  ArgumentData = Mid(OptimCell.Formula, InStr(OptimCell.Formula, "(") + 1, Len(OptimCell.Formula) - InStr(OptimCell.Formula, "(") - 1)
  FormulaParameters = Split(ArgumentData, ",")
  ' Initialize the number of parameters we think we have.  This may need to be adjusted if the third parameter has commas.
  ParameterCount = UBound(FormulaParameters) - LBound(FormulaParameters) + 1
  ' Reassemble the third parameter if there are commas in it...
  For i = 3 To UBound(FormulaParameters)
    FormulaParameters(LBound(FormulaParameters) + 2) = FormulaParameters(LBound(FormulaParameters) + 2) & "," & FormulaParameters(i)
    ' This token is part of the third parameter, so we need to reduce the number of parameters we thought we had.
    ParameterCount = ParameterCount - 1
  Next
  ' At this point ParameterCount should be correct.
  ' Check for a laundry list of errors and set the error message for the first one we encounter...
  If ParameterCount = 1 Then
    ErrorMessage = "Error.  There is one parameter, but " & FunctionName & "() must have 3 parameters."
  ElseIf ParameterCount <> 3 Then
    ErrorMessage = "Error.  There are " & ParameterCount & " parameters, but " & FunctionName & "() must have 3 parameters."
  ElseIf VerifySingleCellSyntax(OptimCell.Worksheet, FormulaParameters(LBound(FormulaParameters))) <> "" Then
    ErrorMessage = "Error.  The first parameter """ & FormulaParameters(LBound(FormulaParameters)) & """" & VerifySingleCellSyntax(OptimCell.Worksheet, FormulaParameters(LBound(FormulaParameters)))
  ElseIf VerifySingleCellSyntax(OptimCell.Worksheet, FormulaParameters(LBound(FormulaParameters) + 1)) <> "" Then
    ErrorMessage = "Error.  The second parameter """ & FormulaParameters(LBound(FormulaParameters) + 1) & """" & VerifySingleCellSyntax(OptimCell.Worksheet, FormulaParameters(LBound(FormulaParameters) + 1))
  ElseIf IsError(OptimCell.Worksheet.Evaluate(FormulaParameters(LBound(FormulaParameters) + 2))) Then
    ErrorMessage = "Error.  The third parameter """ & FormulaParameters(LBound(FormulaParameters) + 2) & """ is an error or references an error."
  ElseIf Not IsNumeric(FormulaParameters(LBound(FormulaParameters) + 2)) Then
    ' Sometimes Excel can't evaluate a simple numerical value like "1".  Don't get me started!
    If IsError(Application.WorksheetFunction.IsError(OptimCell.Worksheet.Evaluate(FormulaParameters(LBound(FormulaParameters) + 2)))) Then
      ErrorMessage = "Error.  The third parameter """ & FormulaParameters(LBound(FormulaParameters) + 2) & """ is an error or references an error."
    End If
  Else
    ' If we got this far, the function was fine.
    ErrorMessage = ""
  End If
  CheckSyntax = ErrorMessage
End Function


' There isn't an elegant way to do this.  Spaghetti code it is!
Function VerifySingleCellSyntax(OptimCellWorksheet As Worksheet, Parameter As String) As String
  Dim DummyCell As Range
  Dim ErrorMessage As String
  ' Assume everything is ok.
  ErrorMessage = ""
  ' Can Excel evaluate this at all?
  On Error GoTo ErrorCannotEvaluate
  Call OptimCellWorksheet.Evaluate(Parameter)
  ' Is this a valid range?
  On Error GoTo ErrorNotRange
  Set DummyCell = OptimCellWorksheet.Evaluate(Parameter)
  ' Does it evaluate to an error?
  If Application.WorksheetFunction.IsError(DummyCell) Then
    GoTo ErrorReferencesErrorCell
  End If
  ' Does this range contain only one cell?
  If DummyCell.Cells.Count > 1 Then
    ErrorMessage = " cannot reference more than one cell at a time."
  End If
  ' Is this cell's value numeric?
  If Not IsNumeric(DummyCell.Value) Then
    ErrorMessage = " is not a valid numerical value."
  End If
  GoTo SyntaxCheckFinish
  ' Let the error handling begin!
ErrorCannotEvaluate:
  ' Cannot evaluate at all.
  ErrorMessage = " is not valid Excel syntax or references an error."
  GoTo SyntaxCheckFinish
ErrorNotRange:
  ' This is not a valid reference to a range.
  ErrorMessage = " does not specify a valid range containing only one cell."
  GoTo SyntaxCheckFinish
ErrorReferencesErrorCell:
  ' This references a cell with an error.
  ErrorMessage = " references a cell with an error."
' We're done with error handling.
SyntaxCheckFinish:
  On Error GoTo 0
  VerifySingleCellSyntax = ErrorMessage
End Function


' The SetTimer API call tells Windows to call this function to set the application timer that initiates the optimization routines.
' It's not like Excel object model workarounds are complicated or anything.
' Unofficial documentation is here: http://www.thatcomicthing.com/comic/15.html
Sub ScheduleOptimizationServiceProc(ByVal HWnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTimer As Long)
  ' This is supposed to be a one-shot timer, so make sure it isn't fired again.
  Call KillTimer(0&, TimerID)
  ' Use the Excel Application's timer so that the service is called when Excel is ready for it.
  Call Application.OnTime(Now, "OptimizationService")
End Sub


' Ensure the list of optimization cells, etc. is up to date.
' Iterate through each optimization cell and optimize one at a time.  Then check for anything that needs to be redone.
' This subroutine decides what needs to be optimized.
Sub OptimizationService()
  Dim LastCalculationStyle As Variant
  Dim LastActiveSheet As Worksheet
  Dim OptimCell As Range
  Dim ArgumentData As String
  Dim Arguments() As String
  Dim FormulaParameters() As String
  Dim TargetCell As Range
  Dim ByChangingCell As Range
  Dim InitialGuess As Double
  Dim Time1 As Currency
  Dim Time2 As Currency
  Dim i As Long
  Dim j As Long
  Dim FunctionName As String
  Dim RecheckIterationCount As Long
  Dim DiscrepancyFound As Boolean
  Dim PrecedentCells() As Range
  Const RecheckMaxIterations = 10
  If ShowDebugText Then
    Debug.Print "The minimization service was started at " & Now & "."
  End If
  ' Kill the timer just to be sure.  (This is useful if OptimizationService is executed independently of a properly scheduled timer.)
  Call KillTimer(0, TimerID)
  ' Clean up the "queue" of Optimize cells to ensure that we don't try to optimize anything that has been deleted.
  Call CleanUpRangeQueue
  ' Is there anything to optimize?
  If IsInitialized(RangeQueue) Then
    ' In case something goes wrong, don't leave the user stranded without automatic calculation or worksheet-level event handling.
    On Error GoTo RestoreCalculationBehavior
    ' Remember the calculation style and prevent calculations that do not pertain to the minimization.
    LastCalculationStyle = Application.Calculation
    Application.Calculation = xlCalculationManual
    ' Speed things up during the calculations.
    Application.ScreenUpdating = False
    ' Don't let anything interrupt this.
    Application.EnableEvents = False
    ' When this is all over, we need to remember what was selected and which sheet it was on...
    Set LastActiveSheet = Application.ActiveSheet
    ' We'll be keeping track of one selection per worksheet.
    ReDim WorkbookSelections(1 To ThisWorkbook.Sheets.Count)
    ' Get the selection from each worksheet so that we can restore selections later.
    For i = 1 To ThisWorkbook.Sheets.Count
      Call ThisWorkbook.Sheets(i).Activate
      Set WorkbookSelections(i) = Application.Selection
    Next
    ' Iterate once for each thing to optimize.
    For i = LBound(RangeQueue, 2) To UBound(RangeQueue, 2)
      ' This is the cell with the optimization function that pertains to this iteration.
      Set OptimCell = RangeQueue(1, i)
      ' Figure out the three function arguments by parsing the text in the cell formula...
      ArgumentData = Mid(OptimCell.Formula, InStr(OptimCell.Formula, "(") + 1, Len(OptimCell.Formula) - InStr(OptimCell.Formula, "(") - 1)
      FormulaParameters = Split(ArgumentData, ",")
      ' Reassemble the third parameter if there are commas in it...
      For j = 3 To UBound(FormulaParameters)
        FormulaParameters(LBound(FormulaParameters) + 2) = FormulaParameters(LBound(FormulaParameters) + 2) & "," & FormulaParameters(j)
      Next
      ' Is this Minimize or MinimizeNonzero?
      FunctionName = Mid(OptimCell.Formula, 2, InStr(OptimCell.Formula, "(") - 2)
      ' Use the arguments to form a reference to the cell...
      ' The Excel Evaluate method uses the active worksheet to evaluate the locations of cells.
      ' It has no other way to know which worksheet they are from unless the worksheet is explicitly stated in the formula (which it probably isn't).
      Set TargetCell = OptimCell.Worksheet.Evaluate(FormulaParameters(LBound(FormulaParameters)))
      Set ByChangingCell = OptimCell.Worksheet.Evaluate(FormulaParameters(LBound(FormulaParameters) + 1))
      ' Before we go any further, ensure that there aren't any errors in the function parameters.
      If Not (Application.WorksheetFunction.IsError(TargetCell) Or Application.WorksheetFunction.IsError(ByChangingCell)) Then
        ' Sometimes Excel's Evaluate() method doesn't work.
        If IsNumeric(FormulaParameters(LBound(FormulaParameters) + 2)) Then
          InitialGuess = CDbl(FormulaParameters(LBound(FormulaParameters) + 2))
          ' Don't let the initial guess violate the constraint, if there is any...
          If LCase(FunctionName) = "minimizenonnegative" Then
            InitialGuess = Max(0, InitialGuess)
          End If
        Else
          ' The user needs to specify an initial guess for the optimization routine.  The initial guess may be a formula.
          InitialGuess = CDbl(OptimCell.Worksheet.Evaluate(FormulaParameters(LBound(FormulaParameters) + 2)))
          ' Don't let the initial guess violate the constraint, if there is any...
          If LCase(FunctionName) = "minimizenonnegative" Then
            InitialGuess = Max(0, InitialGuess)
          End If
        End If
        ' Did the initial guess change?  If so, we assume the user intends to re-optimize this one.
        If RangeQueue(3, i) <> InitialGuess Then
          ' Remember this initial guess for next time.  This may or may not be the first time the initial guess is recorded in RangeQueue for this cell.
          RangeQueue(3, i) = InitialGuess
          ' This will force the cell to be re-optimized using the new initial guess.
          ByChangingCell.Value = InitialGuess
        End If
        ' Check the calculation mode to see if we need to search for precedents.
        If CalculationMode = "ByPrecedents" Then
          ' Where are the precedents that need to be recalculated for the merit function?
          PrecedentCells = ArrangePrecedents(GetAllPrecedents(TargetCell))
        End If
        ' Call the optimization routine and time its performance...
        Call QueryPerformanceCounter(Time1)
        Call PerformOptimization(OptimCell, TargetCell, ByChangingCell, InitialGuess, PrecedentCells, LCase(FunctionName) = "minimizenonnegative")
        Call QueryPerformanceCounter(Time2)
        If ShowDebugText Then
          Debug.Print DisplayElapsedTime("  Total Time to minimize '" & OptimCell.Worksheet.Name & "'!" & OptimCell.Address, Time1, Time2) & " seconds"
        End If
        ' Remember the value of the merit function to check later.
        RangeQueue(2, i) = TargetCell.Value
      End If
    Next
    ' Check all of the merit function values just in case something changed and is no longer optimized.
    ' For example, this can occur when the optimization functions reference each other...
    ' Assume that we need to check for discrepancies...
    DiscrepancyFound = True
    RecheckIterationCount = 0
    ' Loop as long as at least one discrepancy was found and as long as we haven't exceeded the limit on the number of these re-checks.
    Do While DiscrepancyFound And (RecheckIterationCount < RecheckMaxIterations)
      ' We are starting a new iteration, but no discrepancies were found... yet...
      RecheckIterationCount = RecheckIterationCount + 1
      DiscrepancyFound = False
      ' Loop once for each optimization cell.
      For i = LBound(RangeQueue, 2) To UBound(RangeQueue, 2)
        ' This is the cell with the optimization function that pertains to this iteration.
        Set OptimCell = RangeQueue(1, i)
        ' Figure out the three function arguments by parsing the text in the cell formula...
        ArgumentData = Mid(OptimCell.Formula, InStr(OptimCell.Formula, "(") + 1, Len(OptimCell.Formula) - InStr(OptimCell.Formula, "(") - 1)
        FormulaParameters = Split(ArgumentData, ",")
        ' Reassemble the third parameter if there are commas in it...
        For j = 3 To UBound(FormulaParameters)
          FormulaParameters(LBound(FormulaParameters) + 2) = FormulaParameters(LBound(FormulaParameters) + 2) & "," & FormulaParameters(j)
        Next
        ' Is this Minimize or MinimizeNonzero?
        FunctionName = Mid(OptimCell.Formula, 2, InStr(OptimCell.Formula, "(") - 2)
        ' The Excel Evaluate method uses the active worksheet to evaluate the locations of cells.
        ' It has no other way to know which worksheet they are from unless the worksheet is explicitly stated in the formula (which it probably isn't).
        Set TargetCell = OptimCell.Worksheet.Evaluate(FormulaParameters(LBound(FormulaParameters)))
        ' Check the calculation mode to see if we need to search for precedents.
        If CalculationMode = "ByPrecedents" Then
          ' Where are the precedents that need to be recalculated for the merit function?
          PrecedentCells = ArrangePrecedents(GetAllPrecedents(TargetCell))
        End If
        ' Calculation is not automatic, so this cell's merit function needs to be re-evaluated...
        Call CalculateRoutine(PrecedentCells, TargetCell)
        ' At this point we know which cell is the target cell.  We can just check to see if it is still at the last optimum value.
        ' Is it different?
        If TargetCell.Value <> RangeQueue(2, i) Then
          ' There was a discrepancy, so there may be others after this round of optimization.
          DiscrepancyFound = True
          ' The target cell's value is different, so it needs to be reoptimized.
          Set ByChangingCell = OptimCell.Worksheet.Evaluate(FormulaParameters(LBound(FormulaParameters) + 1))
          ' The user needs to specify an initial guess for the optimization routine.  The initial guess may be a formula.
          InitialGuess = CDbl(OptimCell.Worksheet.Evaluate(FormulaParameters(LBound(FormulaParameters) + 2)))
          ' Don't let the initial guess violate the constraint, if there is any...
          If LCase(FunctionName) = "minimizenonnegative" Then
            InitialGuess = Max(0, InitialGuess)
          End If
          ' Optimize the actual cell...
          Call QueryPerformanceCounter(Time1)
          Call PerformOptimization(OptimCell, TargetCell, ByChangingCell, InitialGuess, PrecedentCells, LCase(FunctionName) = "minimizenonnegative")
          Call QueryPerformanceCounter(Time2)
          If ShowDebugText Then
            Debug.Print DisplayElapsedTime("  Total Time to re-minimize '" & OptimCell.Worksheet.Name & "'!" & OptimCell.Address & "'", Time1, Time2) & " seconds"
          End If
          ' Remember the value of the merit function to check later.
          RangeQueue(2, i) = TargetCell.Value
        End If
      Next
    Loop
    ' Go back to the old active worksheet that the user expects to see and restore normal screen updating and event behavior...
RestoreCalculationBehavior:
    ' Be careful putting back the selection...
    For i = 1 To ThisWorkbook.Sheets.Count
      If Not (WorkbookSelections(i) Is Nothing) Then
        If IsObject(WorkbookSelections(i)) Then
          ' If we got this far, there was something valid to select.
          Call ThisWorkbook.Sheets(i).Activate
          Call WorkbookSelections(i).Select
        End If
      End If
    Next
    ' The selections are put back where they belong, but the user is expecting to be on a certain worksheet.
    Call LastActiveSheet.Activate
    Application.Calculation = LastCalculationStyle
    Application.EnableEvents = True
    ' Only update the screen after everything else is ready.
    Application.ScreenUpdating = True
    ' Allow the user to have access to Excel again because we won't be doing any further calculations.
    Application.Interactive = True
    On Error GoTo 0
  End If
  If ShowDebugText Then
    Debug.Print "The minimization service finished at " & Now & "."
  End If
  ' Allow the timer to be set again.
  TimerID = 0
End Sub


' This is the optimization routine.
' OptimCell is the cell that contains the Minimize/MinimizeNonnegative function and references dependent cells.
' TargetCell is the merit function value that will be minimized.
' ByChangingCell is the independent variable that can change.  Its value is unconstrained.
' InitialGuess is the starting value for ByChangingCell, but it is only used if TargetCell is not already at a local minimum.
' The optimization begins with a line search to find suitable starting conditions for a false position method.  The two
' positions Left and Right are updated by quadratic approximations based on points Left and Right and each of their slope values.
' This version of the algorithm is for optimizing a single parameter, which means that all of the matrix math reduces to scalar math.  Ahhh.
Sub PerformOptimization(OptimCell As Range, TargetCell As Range, ByChangingCell As Range, InitialGuess As Double, PrecedentCells() As Range, ConstrainNonnegative As Boolean)
  Dim Delta As Double
  Dim DeltaLeft As Double
  Dim ConstraintActive As Boolean
  Dim LR() As Double
  Dim M() As Double
  Dim LineSearchABC() As Double
  Dim i As Long
  Dim InitialValue As Double
  Dim Derivative As Double
  Dim BackwardsDerivative As Double
  Dim Merit As Double
  Dim MeritNew As Double
  Dim IterationCount As Long
  Dim MaxIterations As Long
  Dim ExitMessage As String
  Dim StabilityValue As Double
  Const StoppingCondition = 0.000000000000001
  ' The merit function hasn't been evaluated yet.
  FunctionEvaluationCount = 0
  ' Allocate space for LR.
  ReDim LR(2, 3)
  ' Specify a differential step for forward-difference numerical derivatives.
  Delta = 0.0001
  ' Initialize this now because it will be used as a dummy variable for a while.
  ReDim M(2, 3)
  ' First we need to test whether the cell's initial value is a local minimum (which most of the time it will be - because the cell was previously optimized).
  InitialValue = ByChangingCell.Value
  ' Use M as a dummy variable for now.
  M(1, 2) = GetYValue(TargetCell, ByChangingCell, PrecedentCells, InitialValue)
  ' Are we constraining x, and would x's reduction by delta violate the nonnegativity constraint?
  If ConstrainNonnegative And (InitialValue - Delta < 0) Then
    ' Are we already at the constraint?
    If InitialValue = 0 Then
      ' Don't bother calculating anything here.
      ' Just put in a zero to trick the optimizer into thinking this is potentially an optimum point (if the right derivative is positive).
      M(1, 3) = 0
      ' The constraint is active.
      ConstraintActive = True
    Else
      ' The x value is not on the constraint, so use the constraint itself as the leftward point.
      ' This is equivalent to reducing Delta to whatever puts x on the constraint, except that it only impacts the left derivative.
      M(1, 3) = (M(1, 2) - GetYValue(TargetCell, ByChangingCell, PrecedentCells, 0)) / InitialValue
    End If
  Else
    ' The constraint, if any, does not apply.  Calculate the derivative on the left side.
    M(1, 3) = (M(1, 2) - GetYValue(TargetCell, ByChangingCell, PrecedentCells, InitialValue - Delta)) / Delta
  End If
  ' For now the derivative of the starting position from the right side is stored in M(2, 3).
  M(2, 3) = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, InitialValue + Delta) - M(1, 2)) / Delta
  
  ' Check the derivative step size and adjust it if necessary...
  ' If the constraint is active, the left side is ambiguous and the right side determines what to do.
  ' If the constraint is not active, both sides determine what to do.
  If (ConstraintActive Or (M(1, 3) = 0)) And (M(2, 3) = 0) Then
    ' Both derivatives were zero, which can't be right.  We'll try a larger step size, which will change later.
    Delta = 0.1
    ' Are we constraining x, and would x's reduction by delta violate the nonnegativity constraint?
    If ConstrainNonnegative And (InitialValue - Delta < 0) Then
      ' Are we already at the constraint?
      If InitialValue = 0 Then
        ' Don't bother calculating anything here.
        ' Just put in a zero to trick the optimizer into thinking this is an optimum point (if the right derivative is positive).
        M(1, 3) = 0
        ' The constraint is active.
        ConstraintActive = True
      Else
        ' The x value is not on the constraint, so use the constraint itself as the leftward point.
        ' This is equivalent to reducing Delta to whatever puts x on the constraint, except that it only impacts the left derivative.
        M(1, 3) = (M(1, 2) - GetYValue(TargetCell, ByChangingCell, PrecedentCells, 0)) / InitialValue
      End If
    Else
      ' The constraint, if any, does not apply.  Calculate the derivative on the left side.
      M(1, 3) = (M(1, 2) - GetYValue(TargetCell, ByChangingCell, PrecedentCells, InitialValue - Delta)) / Delta
    End If
    ' Recalculate the derivative on the right side.
    M(2, 3) = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, InitialValue + Delta) - M(1, 2)) / Delta
  End If
  ' At this point we're assuming that, if it was necessary, the larger step size worked.
  ' If, at the initial guess, the leftward derivative is negative or zero and the rightward derivative is positive, then this must be a local minimum.
  If (M(1, 3) <= 0) And (M(2, 3) > 0) Then
    ' Aha!  The cell is already at a local minimum (or a constrained minimum).
    ' Don't bother optimizing anything.  Just set ByChangingCell to its optimum value and recalculate.
    ByChangingCell = InitialValue
    Call CalculateRoutine(PrecedentCells, TargetCell)
    ' We're done.
    ExitMessage = "  " & OptimCell.Worksheet.Name & "'!" & OptimCell.Address & " is already at a local minimum."
  Else
    ' We're going to start using the initial guess instead of the initial value, so the constraint may not be controlling.
    ConstraintActive = False
    ' We were not already at a local minimum.  Time to begin optimizing...
    ' Everything needs to begin from the initial guess (not necessarily the initial value), so first we will get its rightward derivative...
    M(2, 2) = GetYValue(TargetCell, ByChangingCell, PrecedentCells, InitialGuess)
    M(2, 3) = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, InitialGuess + Delta) - M(2, 2)) / Delta
    ' Then get its leftward derivative.
    M(1, 2) = M(2, 2)
    ' Are we constraining x, and would x's reduction by delta violate the nonnegativity constraint?
    If ConstrainNonnegative And (InitialGuess - Delta < 0) Then
      ' Are we already at the constraint?
      If InitialGuess = 0 Then
        ' Don't bother calculating anything here.
        ' Just put in a negative value to trick the line search to looking to the right (if there is a line search)
        ' or to trick the optimizer into thinking this is an optimum point (if the right derivative is positive).
        M(1, 3) = 0
        ' The constraint is active.
        ConstraintActive = True
      Else
        ' The x value is not on the constraint, so use the constraint itself as the leftward point.
        ' This is equivalent to reducing Delta to whatever puts x on the constraint, except that it only impacts the left derivative.
        M(1, 3) = (M(1, 2) - GetYValue(TargetCell, ByChangingCell, PrecedentCells, 0)) / InitialGuess
      End If
    Else
      ' The constraint, if any, does not apply.  Calculate the derivative on the left side.
      M(1, 3) = (M(1, 2) - GetYValue(TargetCell, ByChangingCell, PrecedentCells, InitialGuess - Delta)) / Delta
    End If
    ' Check the derivatives again just in case.
    If (M(1, 3) <= 0) And (M(2, 3) > 0) Then
      ' Apparently the initial guess was a very good guess.
      ' Set the values appropriately and skip the optimization...
      LR(1, 1) = InitialGuess
      LR(2, 1) = InitialGuess
      ExitMessage = "  Not optimizing '" & OptimCell.Worksheet.Name & "'!" & OptimCell.Address & ".  The initial guess was a good one!"
    Else
      ' Perform a line search that results in valid starting conditions.
      LR = ExponentialLineSearch(TargetCell, ByChangingCell, PrecedentCells, InitialGuess, M(2, 3), ConstrainNonnegative)
      
      ' If LR contains the same point twice, then the x value is optimum at the constraint and further minimization is unnecessary.
      If LR(1, 1) = LR(2, 1) Then
        ' This is as good as it is allowed to get.
        ByChangingCell.Value = LR(1, 1)
        ' Recalculate because we're using a new x value.
        Call CalculateRoutine(PrecedentCells, TargetCell)
        If ShowDebugText Then
          ' Tell the user that we're done here.
          Debug.Print "  " & OptimCell.Worksheet.Name & " '!" & OptimCell.Address & " is bounded by the constraint and cannot be optimized further.  This was determined after " & FunctionEvaluationCount & " function evaluations."
        End If
      Else
        ' At this point LR is initialized with two points (the "Left" point and the "Right" point) and their derivatives.
        ' Also at this point the constraint has been dealt with and is no longer applicable.
        ' There must be at least one local minimum between the Left and Right points because
        ' Left's derivative is negative and Right's derivative is positive.
        ' The rest of this function tries to zero in on the location of the local minimum by iteratively changing LR.
        ' Most of the time it is bringing L and R closer together while maintaining the requirement that the derivative change sign between them.
        ' This is the merit function at this time.  (Just use the lower of the two values.)
        Merit = (LR(1, 2) + LR(2, 2)) / 2
        ' Prepare to iteratively reduce the distance between L and R...
        MaxIterations = 50
        IterationCount = 0
        ' Loop as long as we haven't iterated too many times.
        Do While IterationCount <= MaxIterations
          ' There is the possibility of numerical instability if the two points are too close together.
          ' The numerical instability surfaces in AssembleM during row manipulations of the 3x3 matrix.
          ' If the x coordinates are too close together, they become indistinguishable.
          ' Are the points really, really, really close to each other? (5E-9 of their value)
          If LR(1, 1) <> 0 Then
            ' If the division won't cause an error, use a value related to the numerical precision.
            StabilityValue = Abs(Abs(LR(2, 1) / LR(1, 1)) - 1)
          Else
            ' Otherwise just use the difference.
            StabilityValue = Abs(LR(2, 1) - LR(1, 1))
          End If
          If StabilityValue < 0.000000005 Then
            ' Be prepared to tell the user what happened.
            ExitMessage = "  Stopped optimizing '" & OptimCell.Worksheet.Name & "'!" & OptimCell.Address & " to avoid numerical instability.  The points (" & LR(1, 1) & ", " & LR(2, 1) & ") are sufficiently close together.  The exit value is: " & StabilityValue & " < 5E-9"
            ' Stop everything!  This is as close as we'll get.
            MaxIterations = IterationCount
          Else
            ' Calculate the values for two midpoints M.  These are potentially better points.
            M = QAssembleM(TargetCell, ByChangingCell, PrecedentCells, LR, Delta)
            ' Update LR using the four known points: Two (old) points in LR and two (new) points in M.
            LR = QUpdateLR(LR, M)
            ' We have a new merit function.
            MeritNew = (LR(1, 2) + LR(2, 2)) / 2
            ' Is the improvement small enough that we can say we're done?
            ' Test for the stopping condition, which is that the merit function is barely changing.
            If (Merit - MeritNew < StoppingCondition) Then
              ' Trick the loop into stopping with this iteration.
              MaxIterations = IterationCount
            End If
            ' Remember this merit function for the next iteration.
            Merit = MeritNew
            ' Use a better Delta on the next iteration.  The differential is 1E-6 of the difference in x between the points or 1E-10.
            ' One of the perks of using this kind of optimization scheme (where we converge on the minimum from opposite sides) is
            ' how easily and confidently we can pick a good differential step size Delta.
            Delta = Max(Abs(LR(2, 1) - LR(1, 1)) / 1000000, 0.0000000001)
          End If
          ' We just finished one more iteration.
          IterationCount = IterationCount + 1
        Loop
      End If
    End If
    ' This provides the best value for the merit function given that both sides converge to the same value.
    ' Set the worksheet cell equal to this value.
    ByChangingCell.Value = (LR(1, 1) + LR(2, 1)) / 2
    Call CalculateRoutine(PrecedentCells, TargetCell)
    ' There is a small chance that the merit function was better at Left or Right, so we need to verify that there isn't a better value at one of those points...
    ' Was Left better?
    If LR(1, 2) < TargetCell.Value Then
      ' Change to Left.
      ByChangingCell.Value = LR(1, 1)
      ' Recalculate because we're using a new x value.
      Call CalculateRoutine(PrecedentCells, TargetCell)
    End If
    ' We also need to check Right.
    If LR(2, 2) < TargetCell.Value Then
      ' It was better here.  Use this value.
      ByChangingCell.Value = LR(2, 1)
      ' Recalculate because we're using a new x value.
      Call CalculateRoutine(PrecedentCells, TargetCell)
    End If
    If ShowDebugText Then
      If IterationCount = 1 Then
        Debug.Print "  --> Merit function for '" & OptimCell.Worksheet.Name & "'!" & OptimCell.Address & " optimized in " & IterationCount & " iteration and " & FunctionEvaluationCount & " function evaluations."
      Else
        Debug.Print "  --> Merit function for '" & OptimCell.Worksheet.Name & "'!" & OptimCell.Address & " optimized in " & IterationCount & " iterations and " & FunctionEvaluationCount & " function evaluations."
      End If
    End If
  End If
  ' Display additional information if there is any.
  If ShowDebugText And (ExitMessage <> "") Then
    Debug.Print ExitMessage
  End If
End Sub


' Recalculate cells based on the appropriate calculation method.
Sub CalculateRoutine(PrecedentCells() As Range, TargetCell As Range)
  If CalculationMode = "ByApplication" Then
    Call Application.Calculate
    ' We just evaluated the merit function again.
    FunctionEvaluationCount = FunctionEvaluationCount + 1
  ElseIf CalculationMode = "ByPrecedents" Then
    Call RecalculateRanges(PrecedentCells)
    Call TargetCell.Calculate
    ' We just evaluated the merit function again.
    FunctionEvaluationCount = FunctionEvaluationCount + 1
  ElseIf CalculationMode = "ByNaivePrecedents" Then
    ' This only accesses precedents if they are all located on the same worksheet.  Very risky!
    Call TargetCell.Precedents.Calculate
    Call TargetCell.Calculate
    ' We just evaluated the merit function again.
    FunctionEvaluationCount = FunctionEvaluationCount + 1
  Else
    Debug.Print "Error.  The constant CalculationMode is not set correctly.  Defaulting to ByApplication"
    Call Application.Calculate
    ' We just evaluated the merit function again.
    FunctionEvaluationCount = FunctionEvaluationCount + 1
  End If
End Sub


' Returns the address of a range Range1 using the minimum amount of information to identify the range relative to range RelativeRange.
' Possible formatting is:
'   A1
'   'Sheet2'!A1
'   [WorkbookName.xlsm]Sheet2'!A1
Function FormatRangeAddress(Range1 As Range, RelativeRange As Range) As String
  Dim SheetName1 As String
  Dim SheetName2 As String
  Dim WorkbookName1 As String
  Dim WorkbookName2 As String
  ' Isolate the sheet names...
  ' Start with the whole address for the cell...
  SheetName1 = Range1.Address(False, False, xlA1, True)
  SheetName2 = RelativeRange.Address(False, False, xlA1, True)
  ' Strip away the workbook name...
  SheetName1 = Mid(SheetName1, InStr(SheetName1, "]") + 1, 999)
  SheetName2 = Mid(SheetName2, InStr(SheetName2, "]") + 1, 999)
  ' If there is a single quote, strip away the single quote and everything after it....
  If InStr(SheetName1, "'") <> 0 Then
    SheetName1 = Mid(SheetName1, 1, InStr(SheetName1, "'") - 1)
  Else
    ' There is no single quote, so remove the exclamation point and everything after it.
    SheetName1 = Mid(SheetName1, 1, InStr(SheetName1, "!") - 1)
  End If
  If InStr(SheetName2, "'") Then
    SheetName2 = Mid(SheetName2, 1, InStr(SheetName2, "'") - 1)
  Else
    ' There is no single quote, so remove the exclamation point and everything after it.
    SheetName2 = Mid(SheetName2, 1, InStr(SheetName2, "!") - 1)
  End If
  ' Isolate the workbook names...
  ' Start with the whole address for the cell...
  WorkbookName1 = Range1.Address(False, False, xlA1, True)
  WorkbookName2 = RelativeRange.Address(False, False, xlA1, True)
  ' The workbook name is everything inside right-angle brackets.  Keep the brackets...
  WorkbookName1 = Mid(WorkbookName1, InStr(WorkbookName1, "["), InStr(WorkbookName1, "]") - InStr(WorkbookName1, "[") + 1)
  WorkbookName2 = Mid(WorkbookName2, InStr(WorkbookName2, "["), InStr(WorkbookName2, "]") - InStr(WorkbookName2, "[") + 1)
  ' Are the ranges in the same workbook?
  If WorkbookName1 = WorkbookName2 Then
    ' The only thing left to check is if the ranges are on the same sheet.
    If SheetName1 = SheetName2 Then
      ' They are in the same workbook and on the same sheet, so just use simple formatting.
      FormatRangeAddress = Range1.Address(False, False, xlA1, False)
    Else
      ' They are in the same workbook but on a different sheet, so show the sheet name and the address.
      FormatRangeAddress = "'" & SheetName1 & "'!" & Range1.Address(False, False, xlA1, False)
    End If
  Else
    ' They are in different places entirely, so just show the whole address.
    FormatRangeAddress = Range1.Address(False, False, xlA1, True)
  End If
End Function


' Add this range to the list of ranges that potentially contain calls to Minimize.
' The return value is the index of R in RangeQueue.
Function AddRange(r As Range) As Long
  Dim RQueueIndex As Long
  ' If the array is initialized, enlarge it.  Otherwise initialize it...
  If IsInitialized(RangeQueue) Then
    ' Where is the range in the list of ranges?
    RQueueIndex = QueueIndex(r)
    ' Is it not already there?
    If RQueueIndex = -1 Then
      ' Add it to the Range queue.
      ReDim Preserve RangeQueue(3, UBound(RangeQueue, 2) + 1)
      ' Add this range at the end...
      RQueueIndex = UBound(RangeQueue, 2)
      Set RangeQueue(1, RQueueIndex) = r
      RangeQueue(2, RQueueIndex) = -99999
      ' The initial guess is irrelevant at this time because it will be populated right before the cell is optimized for the first time.
      RangeQueue(3, RQueueIndex) = -99999
    End If
  Else
    ' Make room for only one item.
    ReDim RangeQueue(3, 1)
    ' Add this range at the end...
    RQueueIndex = UBound(RangeQueue, 2)
    Set RangeQueue(1, RQueueIndex) = r
    ' The merit function value is irrelevant at this time because it will be populated right before the cell is optimized for the first time.
    RangeQueue(2, RQueueIndex) = -99999
    ' Ditto for the initial guess.
    RangeQueue(3, RQueueIndex) = -99999
  End If
  AddRange = RQueueIndex
End Function


' This function was shamelessly copied from:
' http://stackoverflow.com/questions/12127311/vba-what-happens-to-range-objects-if-user-deletes-cells
' Apparently there is no elegant way to do this either.
Function RangeWasDeclaredAndEntirelyDeleted(r As Variant) As Boolean
  Dim TestAddress As String
  If r Is Nothing Then
      Exit Function
  End If
  On Error Resume Next
  TestAddress = r.Address
  If Err.Number = 424 Then    'object required
    RangeWasDeclaredAndEntirelyDeleted = True
  End If
End Function


' The list of cells that potentially contain calls to Minimize needs to be kept up to date.
Sub CleanUpRangeQueue()
  Dim RangeAddresses As String
  Dim OptimCell As Range
  Dim OriginalCellCount As Long
  Dim UpdatedCellCount As Long
  Dim NewRange As Range
  Dim i
  Dim NewArray() As Variant
  ' Is there a list of cells?
  If IsInitialized(RangeQueue) Then
    ' We haven't counted the updated cells yet.
    UpdatedCellCount = 0
    ' How many cells can there potentially be?
    OriginalCellCount = UBound(RangeQueue, 2) - LBound(RangeQueue, 2) + 1
    ' Iterate once for each potential cell with an Optimize formula.
    For i = LBound(RangeQueue, 2) To UBound(RangeQueue, 2)
      ' Does this cell still exist?
      If Not RangeWasDeclaredAndEntirelyDeleted(RangeQueue(1, i)) Then
        ' Is this still a valid cell to optimize?
        If (LCase(Mid(RangeQueue(1, i).Formula, 1, 10)) = "=minimize(") Or (LCase(Mid(RangeQueue(1, i).Formula, 1, 21)) = "=minimizenonnegative(") Then
          ' Count this one.
          UpdatedCellCount = UpdatedCellCount + 1
        End If
      End If
    Next
    ' Do we need to make any changes?
    If UpdatedCellCount <> OriginalCellCount Then
      ' Are there any valid ranges at all?
      If UpdatedCellCount = 0 Then
        ' Just put in one element (because we can't uninitialize the array).
        ' One element is fairly harmless.
        ReDim Preserve RangeQueue(2, 1)
      Else
        ' It looks like there was one or more cells with calls to Minimize, so let's allocate space and keep track of them.
        ReDim NewArray(3, UpdatedCellCount)
        ' We haven't updated any cells in the new array.
        UpdatedCellCount = 0
        ' Loop once for each potentially active optimization cell.
        For i = LBound(RangeQueue, 2) To UBound(RangeQueue, 2)
          ' Does this cell still exist?
          If Not RangeWasDeclaredAndEntirelyDeleted(RangeQueue(1, i)) Then
            ' Is this a valid cell to optimize?
            If (LCase(Mid(RangeQueue(1, i).Formula, 1, 10)) = "=minimize(") Or (LCase(Mid(RangeQueue(1, i).Formula, 1, 21)) = "=minimizenonnegative(") Then
              ' We are about to add another range.
              UpdatedCellCount = UpdatedCellCount + 1
              ' Include this cell and its merit function value and its initial value...
              Set NewArray(1, UpdatedCellCount) = RangeQueue(1, i)
              NewArray(2, UpdatedCellCount) = RangeQueue(2, i)
              NewArray(3, UpdatedCellCount) = RangeQueue(3, i)
            End If
          End If
        Next
        ' Use the new array in place of the old one.
        RangeQueue = NewArray
      End If
    End If
  End If
End Sub


' Is this range in the array?  Returns the index number or -1.
Function QueueIndex(r As Variant)
  Dim i As Long
  ' Assume it wasn't found.
  QueueIndex = -1
  ' Loop once for each item in the array.
  For i = LBound(RangeQueue, 2) To UBound(RangeQueue, 2)
    ' Make sure that the array actually contains an object before attempting to read its Address property.
    If IsObject(RangeQueue(1, i)) Then
      If Not RangeWasDeclaredAndEntirelyDeleted(RangeQueue(1, i)) Then
        ' Is this a match?  Use the external address to distinguish between cells with the same address on different worksheets.
        If RangeQueue(1, i).Address(True, True, xlA1, True) = r.Address(True, True, xlA1, True) Then
          ' This is a match.  Don't bother looking for more...
          QueueIndex = i
          Exit For
        End If
      End If
    End If
  Next
End Function


' Returns the lower of the two values in the second column of a 2x2 array.
Function MinMerit(LR() As Double) As Double()
  Dim Result(2) As Double
  Result(1) = 1
  Result(2) = LR(1, 2)
  If LR(2, 2) < LR(1, 2) Then
    Result(1) = 2
    Result(2) = LR(2, 2)
  End If
  MinMerit = Result
End Function


' Chooses new LR points based on LR and M.
Function QUpdateLR(LR() As Double, M() As Double) As Double()
  Dim FourPoints(4, 3) As Double
  Dim Temp(3) As Double
  Dim i As Long
  Dim j As Long
  Dim MinValue As Double
  Dim MinIndex As Long
  Dim ComplementaryIndex As Long
  
  ' Assemble the four candidate points into a single array...
  ' Add the LR points, Left and Right...
  For i = 1 To 2
    For j = LBound(FourPoints, 2) To UBound(FourPoints, 2)
      FourPoints(i, j) = LR(i, j)
    Next
  Next
  ' Add two M points, M1 and M2...
  For i = 1 To 2
    For j = LBound(FourPoints, 2) To UBound(FourPoints, 2)
      FourPoints(i + 2, j) = M(i, j)
    Next
  Next
  ' Sort the four points in order of ascending x coordinate using a BubbleSort algorithm...
  For i = LBound(FourPoints, 1) To UBound(FourPoints, 1) - 1
    For j = i + 1 To UBound(FourPoints, 1)
      ' Is the second (j) point further to the left (than the i point)?
      If FourPoints(j, 1) < FourPoints(i, 1) Then
        ' We need to swap the points at Indexes i and j...
        Temp(1) = FourPoints(j, 1)
        Temp(2) = FourPoints(j, 2)
        Temp(3) = FourPoints(j, 3)
        FourPoints(j, 1) = FourPoints(i, 1)
        FourPoints(j, 2) = FourPoints(i, 2)
        FourPoints(j, 3) = FourPoints(i, 3)
        FourPoints(i, 1) = Temp(1)
        FourPoints(i, 2) = Temp(2)
        FourPoints(i, 3) = Temp(3)
      End If
    Next
  Next
  ' At this point the four points are sorted in order by ascending x coordinate.
  ' Now we need to know which one has the best merit function value...
  ' Assume the first item is the minimum...
  MinValue = FourPoints(1, 2)
  MinIndex = LBound(FourPoints, 1)
  ' Look through the other three points to see if one of them is lower.
  For i = LBound(FourPoints, 1) + 1 To UBound(FourPoints)
    If FourPoints(i, 2) < MinValue Then
      ' This is the new minimum value.
      MinValue = FourPoints(i, 2)
      ' Remember this index.
      MinIndex = i
    End If
  Next
  ' We haven't found the point on the other side of Left or Right yet, so assume we didn't find one.
  ComplementaryIndex = -1
  ' LR must contain the best point (with the lowest y value).  The question now is whether it is a Left point or a Right point.
  If FourPoints(MinIndex, 3) < 0 Then
    ' The derivative is negative, so this must be a Left point.
    LR(1, 1) = FourPoints(MinIndex, 1)
    LR(1, 2) = FourPoints(MinIndex, 2)
    LR(1, 3) = FourPoints(MinIndex, 3)
    ' Now that we have the Left point, we need to find a Right point.
    ' Search for a valid point to the right of this one.
    For i = MinIndex + 1 To UBound(FourPoints, 1)
      ' Does this point have a positive derivative?
      If FourPoints(i, 3) > 0 Then
        ' This one will do.
        ComplementaryIndex = i
        ' Copy it to Right...
        LR(2, 1) = FourPoints(ComplementaryIndex, 1)
        LR(2, 2) = FourPoints(ComplementaryIndex, 2)
        LR(2, 3) = FourPoints(ComplementaryIndex, 3)
        ' Don't bother searching for more of these.
        Exit For
      End If
    Next
  Else
    ' The derivative is positive (or zero), so this must be a Right point (or a local minimum).
    LR(2, 1) = FourPoints(MinIndex, 1)
    LR(2, 2) = FourPoints(MinIndex, 2)
    LR(2, 3) = FourPoints(MinIndex, 3)
    ' Now that we have the Right point, we need to find a Left point.
    ' Search for a valid point to the left of this one.
    For i = MinIndex - 1 To LBound(FourPoints, 1) Step -1
      ' Does this point have a negative derivative?
      If FourPoints(i, 3) < 0 Then
        ' This one will do.
        ComplementaryIndex = i
        ' Copy it to Left...
        LR(1, 1) = FourPoints(ComplementaryIndex, 1)
        LR(1, 2) = FourPoints(ComplementaryIndex, 2)
        LR(1, 3) = FourPoints(ComplementaryIndex, 3)
        ' Don't bother searching for more of these.
        Exit For
      End If
    Next
  End If
  ' At this point we have new, valid LR values.
  QUpdateLR = LR
End Function


' Returns M, which is an array (x, y, and dy/dx) for two points.
' The two points have x values at the vertexes of fitted parabolas.
' y and dy/dx are determined by evaluating the merit function.
Function QAssembleM(TargetCell As Range, ByChangingCell As Range, PrecedentCells() As Range, LR() As Double, Delta As Double) As Double()
  Dim M(2, 3) As Double
  Dim A1(3, 4) As Double
  Dim A2(3, 4) As Double
  Dim i As Long
  Dim j As Long
  Dim Ratio As Double
  Dim ABC() As Double
  ' Model the parabola with the slope at L...
  ' Populate the values of A...
  ' Start with the derivative information because that needs to be in the first row...
  ' The derivative dy/dx = 2*A*x + b.
  A1(1, 1) = 2 * LR(1, 1)
  A1(1, 2) = 1
  A1(1, 3) = 0
  A1(1, 4) = LR(1, 3)
  ' Add the equality conditions for the left and right sides.
  A1(2, 1) = LR(1, 1) ^ 2
  A1(2, 2) = LR(1, 1)
  A1(2, 3) = 1
  A1(2, 4) = LR(1, 2)
  A1(3, 1) = LR(2, 1) ^ 2
  A1(3, 2) = LR(2, 1)
  A1(3, 3) = 1
  A1(3, 4) = LR(2, 2)
  ' Take a moment to copy the values from A1 to A2, except for the derivative condition in the first row...
  For i = LBound(A1, 1) + 1 To UBound(A1, 1)
    For j = LBound(A1, 2) To UBound(A1, 2)
      A2(i, j) = A1(i, j)
    Next
  Next
  ' Include the derivative condition in the first row of A2...
  ' The derivative dy/dx = 2*A*x + b.
  A2(1, 1) = 2 * LR(2, 1)
  A2(1, 2) = 1
  A2(1, 3) = 0
  A2(1, 4) = LR(2, 3)
  ' Solve for the M1 parabola's parameters A, B, and C.
  ABC = Solve3x3AugmentedMatrix(A1)
  ' Calculate the x coordinate for M1, which is the parabola's vertex.
  M(1, 1) = -ABC(2) / (2 * ABC(1))
  ' Solve for the M2 parabola's parameters A, B, and C.
  ABC = Solve3x3AugmentedMatrix(A2)
  ' Calculate the x coordinate for M2, which is the parabola's vertex.
  M(2, 1) = -ABC(2) / (2 * ABC(1))
  ' At this point we have the x values for the two midpoints M.  Now we need y and dy/dx.
  For i = LBound(M, 1) To UBound(M, 1)
    ' We need the function value first.
    M(i, 2) = GetYValue(TargetCell, ByChangingCell, PrecedentCells, M(i, 1))
    ' Now the derivative...
    M(i, 3) = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, M(i, 1) + Delta) - M(i, 2)) / Delta
  Next
  QAssembleM = M
End Function


' Solve a 3x4 augmented matrix.
' This assumes that element 1, 3 is zero, which is the case for the type of curve fitting we are doing.
' The first row is the derivative condition and the next two rows are the continuity conditions.
' The result is the coefficient vector [A, B, C] for y = A * x ^ 2 + B * x + C
' This is a very rudimentary solver that is tailored to this particular type of curve fitting.
Function Solve3x3AugmentedMatrix(A() As Double) As Double()
  Dim Ratio As Double
  Dim i As Long
  Dim ABC(3) As Double
  
  ' We have to ensure that the first two diagonal elements are nonzero.
  ' Start with the first diagonal.
  If A(1, 1) = 0 Then
    ' Is there a nonzero element in the second row?
    If A(2, 1) <> 0 Then
      ' Swap the first and second rows, using Ratio as a dummy variable...
      For i = LBound(A, 2) To UBound(A, 2)
        Ratio = A(1, i)
        A(1, i) = A(2, i)
        A(2, i) = Ratio
      Next
    Else
      ' The second row didn't work, so we have to assume that the third row will.
      ' Swap the first and third rows, using Ratio as a dummy variable...
      For i = LBound(A, 2) To UBound(A, 2)
        Ratio = A(1, i)
        A(1, i) = A(3, i)
        A(3, i) = Ratio
      Next
    End If
  End If
  ' Check the second diagonal.
  If A(2, 2) = 0 Then
    ' We have to assume that the third row will help with this.
    ' Swap the second and third rows, using Ratio as a dummy variable...
    For i = LBound(A, 2) To UBound(A, 2)
      Ratio = A(2, i)
      A(2, i) = A(3, i)
      A(3, i) = Ratio
    Next
  End If
  
  ' Zero out the first column of the third row if it isn't already zero...
  ' Try to use the second row because Ratio will be close to -1, which works best with the available precision.
  If A(3, 1) <> 0 Then
    Ratio = -A(3, 1) / A(2, 1)
    For i = LBound(A, 2) To UBound(A, 2)
      A(3, i) = A(3, i) + A(2, i) * Ratio
    Next
  End If
  ' Zero out the first column of the second row if it isn't already zero...
  If A(2, 1) <> 0 Then
    Ratio = -A(2, 1) / A(1, 1)
    For i = LBound(A, 2) To UBound(A, 2)
      A(2, i) = A(2, i) + A(1, i) * Ratio
    Next
  End If
  ' Zero out the second column of the third row if it isn't already zero...
  If A(3, 2) <> 0 Then
    Ratio = -A(3, 2) / A(2, 2)
    For i = LBound(A, 2) To UBound(A, 2)
      A(3, i) = A(3, i) + A(2, i) * Ratio
    Next
  End If
  ' Solve for C.
  ABC(3) = A(3, 4) / A(3, 3)
  ' Solve for B.
  A(2, 4) = A(2, 4) - A(2, 3) * ABC(3)
  ABC(2) = A(2, 4) / A(2, 2)
  ' Solve for A.
  A(1, 4) = A(1, 4) - A(1, 3) * ABC(3) - A(1, 2) * ABC(2)
  ABC(1) = A(1, 4) / A(1, 1)
  Solve3x3AugmentedMatrix = ABC
End Function


' Performs a line search in the direction of decreasing merit function value and refines the results of the line search to meet the requirements for starting conditions for this algorithm.
' Each point in the line search is twice as far away as the one before it.  The stopping condition is a point above its predecessor, indicating the presence of a local minimum.
' Returns LR values (points and derivatives) surrounding the local minimum.  These are valid starting positions for the minimization routine.
Function ExponentialLineSearch(TargetCell As Range, ByChangingCell As Range, PrecedentCells() As Range, InitialGuess As Double, InitialGuessDerivative As Double, ConstrainNonnegative As Boolean) As Double()
  Dim ABC(3, 3) As Double
  Dim i As Long
  Const MaxSearchIterations = 41 ' 41 moves the search by FirstIncrementSize * 2 ^ 40
  Dim FirstIncrementSize As Double
  Dim IterationCount As Long
  Dim NextValue As Double
    
  ' Initialize the derivative.
  ABC(2, 3) = InitialGuessDerivative
  ' We will search in the direction of the negative gradient.
  FirstIncrementSize = -0.05 * Sgn(InitialGuessDerivative)
  ' If the search direction is ambiguous, assume we search to the right...
  If FirstIncrementSize = 0 Then
    FirstIncrementSize = 0.05
  End If
  
  ' We haven't searched yet.
  IterationCount = 0
  ' First we'll put InitialGuess in the middle.
  ' Loop once for the three positions.
  For i = LBound(ABC, 1) To UBound(ABC, 1)
    ' This is the next value that we want to evaluate the merit function at.
    NextValue = InitialGuess + FirstIncrementSize * (IterationCount - 1)
    ' This point cannot be allowed to violate the constraint.  Does that look like it will be a problem?
    If ConstrainNonnegative And (NextValue < 0) Then
      ' Use the lowest allowable value of zero.
      ABC(i, 1) = 0
    Else
      ' A constraint violation is not a problem.  Use the value as calculated.
      ABC(i, 1) = NextValue
    End If
    ' Evaluate the point, which at this point is guaranteed to be within the constraint (if any).
    ABC(i, 2) = GetYValue(TargetCell, ByChangingCell, PrecedentCells, ABC(i, 1))
    ' We just used up an increment iteration.
    IterationCount = IterationCount + 1
  Next
  
  ' At this point ABC is initialized with the first three points in the line search.
  ' We are looking for A, B, C values where B is lower than both A and C.
  Do While (ABC(3, 2) <= ABC(2, 2)) And (IterationCount <= MaxSearchIterations)
    ' Shift over the last two sets of values...
    ABC(1, 1) = ABC(2, 1)
    ABC(1, 2) = ABC(2, 2)
    ABC(2, 1) = ABC(3, 1)
    ABC(2, 2) = ABC(3, 2)
    
    NextValue = InitialGuess + FirstIncrementSize * 2 ^ (IterationCount - 2)
    ' Will this value cause a problem?
    If ConstrainNonnegative And NextValue < 0 Then
      ' Use zero by default and stop iterating because further iterations will only make things worse...
      ABC(3, 1) = 0
      ABC(3, 2) = GetYValue(TargetCell, ByChangingCell, PrecedentCells, ABC(3, 1))
      IterationCount = MaxSearchIterations
    Else
      ' Populate the new x and y values for Point C...
      ABC(3, 1) = NextValue
      ABC(3, 2) = GetYValue(TargetCell, ByChangingCell, PrecedentCells, ABC(3, 1))
    End If
    ' We just used up an increment iteration.
    IterationCount = IterationCount + 1
  Loop
  ' At this point ABC contains three points where the derivative changes sign somewhere between the outer two (or an outer point encounters the constraint).
  ' However, that is not sufficient for the type of search we need to perform.
  ' We need to narrow it down to two points with opposite derivatives (indicating a local minimum).
  ExponentialLineSearch = EnforceDerivativeRequirements(TargetCell, ByChangingCell, PrecedentCells, ABC, ConstrainNonnegative)
End Function


' Returns the y value at a given x value using spreadsheet calculations.
' This also increments the global variable FunctionEvaluationCount, which keeps track of how many times the function was evaluated.
Function GetYValue(TargetCell As Range, ByChangingCell As Range, PrecedentCells() As Range, x As Double) As Double
  ' Set the x value we're measuring.
  ByChangingCell.Value = x
  ' Recalculate according to the current calculation settings.
  Call CalculateRoutine(PrecedentCells, TargetCell)
  ' Now we have a new y value that was calculated with updated precedent cells.
  GetYValue = TargetCell.Value
  ' We just evaluated the merit function again.
  FunctionEvaluationCount = FunctionEvaluationCount + 1
End Function


' Refines the first or third point of a line search so that the first or third point has the correct derivative.
' There are two points.  The left point must have a negative derivative and the right point must have a positive derivative.
' The assumption of continunity tells us that the function's derivative must change sign somewhere between the two points.
' However, if the ABC points encounter the constraint (and if that appears to be optimum), the return value contains that constrained optimum only.
Function EnforceDerivativeRequirements(TargetCell As Range, ByChangingCell As Range, PrecedentCells() As Range, ABC() As Double, ConstrainNonnegative As Boolean) As Double()
  Dim Delta As Double
  Dim LeftX As Double
  Dim LeftY As Double
  Dim RightX As Double
  Dim RightY As Double
  Dim MidX As Double
  Dim MidY As Double
  Dim x0 As Double
  Dim Derivative As Double
  Dim IterationCount As Long
  Dim MaxIterations As Long
  Dim LR(2, 3) As Double
  Const MaxIterationsConst = 30
  
  ' Sort the array ABC.  The middle point is where it should be, but the outer points may be transposed.
  ' Does the last point come before the first point?
  If ABC(3, 1) < ABC(1, 1) Then
    ' Swap the two points using x0 as a dummy variable...
    x0 = ABC(1, 1)
    ABC(1, 1) = ABC(3, 1)
    ABC(3, 1) = x0
    x0 = ABC(1, 2)
    ABC(1, 2) = ABC(3, 2)
    ABC(3, 2) = x0
  End If
  ' If the binary search needs this many iterations, it's hopeless anyway.
  MaxIterations = MaxIterationsConst
  ' Assume that we won't have adjacent points so close together in x.
  Delta = (ABC(3, 1) - ABC(1, 1)) / (3 * 2 ^ (MaxIterations))
  ' The derivative may have been obsolete, so calculate a new one using this Delta value.
  ABC(2, 3) = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, ABC(2, 1) + Delta) - ABC(2, 2)) / Delta
  ' The derivative at the middle point tells us whether we need to refine the first point or the third point.
  ' Henceforth the variable Derivative will represent the derivative at a variety of locations.
  If ABC(2, 3) > 0 Then
    ' The derivative at B is positive, which means we need to refine the left point to ensure its derivative is negative.
    ' Calculate the derivative at the left point.
    Derivative = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, ABC(1, 1) + Delta) - ABC(1, 2)) / Delta
    ' Is the derivative already negative?
    If Derivative < 0 Then
      ' Remember this derivative as a courtesy to the next function that uses ABC.
      ABC(1, 3) = Derivative
      ' Don't do anything else to ABC; it already meets the requirement.
    Else
      ' The derivative was positive, which means it must turn negative somewhere between the left and center points.
      ' Set the boundaries for the forthcoming search...
      LeftX = ABC(1, 1)
      LeftY = ABC(1, 2)
      RightX = ABC(2, 1)
      RightY = ABC(2, 2)
      ' We haven't searched yet.
      IterationCount = 0
      ' Perform a binary search.  During this search, Left, Middle, and Right refer to the boundaries pertaining to the search algorithm.
      Do While IterationCount < MaxIterations
        ' The middle point is halfway between LeftX and RightX.  This is essentially a binary search algorithm...
        MidX = (LeftX + RightX) / 2
        MidY = GetYValue(TargetCell, ByChangingCell, PrecedentCells, MidX)
        Derivative = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, MidX + Delta) - MidY) / Delta
        ' Does the derivative do what we need it to do here?
        If Derivative < 0 Then
          ' The derivative is negative, so we're done searching.  Prevent further iterations.
          MaxIterations = IterationCount
          ' Populate ABC with the new value...
          ABC(1, 1) = MidX
          ABC(1, 2) = MidY
          ABC(1, 3) = Derivative
        Else
          ' The derivative wasn't what we needed, so we have to try a new (smaller) search region...
          ' If the midpoint is above the right, then midpoint is the new left.  (We prefer the right half of the search region.)
          If MidY > RightY Then
            ' Get the new Left values...
            LeftX = MidX
            LeftY = MidY
          Else
            ' If the midpoint is at or below the right, then midpoint is the new right. (We prefer the left half of the search region.)
            ' Get the new Right values...
            RightX = MidX
            RightY = MidY
          End If
        End If
        ' We just finished another iteration.
        IterationCount = IterationCount + 1
      Loop
      ' Check to see if we maxed out our search iterations.
      If (IterationCount >= MaxIterationsConst) Then
        ' This is the derivative at the leftmost point.
        Derivative = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, LeftX + Delta) - LeftY) / Delta
        ' After all that, do we also find ourselves at a leftmost point with a positive derivative?
        If Derivative >= 0 Then
          ' Move the leftmost and rightmost points to the same location (which must be zero) to signify that this problem's solution is constrained.
          ABC(1, 1) = LeftX
          ABC(1, 2) = LeftY
          ABC(2, 1) = ABC(1, 1)
          ABC(2, 2) = ABC(1, 2)
        End If
      End If
    End If
  Else
    ' The derivative at B must have been negative or zero, which means we need to refine the right point to ensure its derivative is positive.
    ' Calculate the derivative at the right point.
    Derivative = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, ABC(3, 1) + Delta) - ABC(3, 2)) / Delta
    ' Is the derivative already positive?
    If Derivative > 0 Then
      ' Remember this derivative as a courtesy to the next function that uses ABC.
      ABC(3, 3) = Derivative
      ' Don't do anything else to ABC; it already meets the requirement.
    Else
      ' The derivative was negative, which means it must turn positive somewhere between the right and center points.
      ' Set the boundaries for the forthcoming search...
      LeftX = ABC(2, 1)
      LeftY = ABC(2, 2)
      RightX = ABC(3, 1)
      RightY = ABC(3, 2)
      ' We haven't searched yet.
      IterationCount = 0
      ' Perform a binary search.  During this search, Left, Middle, and Right refer to the boundaries pertaining to the search algorithm.
      Do While IterationCount < MaxIterations
        ' The middle point is halfway between LeftX and RightX.  This is essentially a binary search algorithm...
        MidX = (LeftX + RightX) / 2
        MidY = GetYValue(TargetCell, ByChangingCell, PrecedentCells, MidX)
        Derivative = (GetYValue(TargetCell, ByChangingCell, PrecedentCells, MidX + Delta) - MidY) / Delta
        ' Does the derivative do what we need it to do here?
        If Derivative > 0 Then
          ' The derivative is positive, so we're done searching.  Prevent further iterations.
          MaxIterations = IterationCount
          ' Populate ABC with the new value...
          ABC(3, 1) = MidX
          ABC(3, 2) = MidY
          ABC(3, 3) = Derivative
        Else
          ' The derivative wasn't what we needed, so we have to try a new (smaller) search region...
          ' If the midpoint is above the left, then midpoint is the new right.  (We prefer the left half of the search region.)
          If MidY > LeftY Then
            ' Get the new Right values...
            RightX = MidX
            RightY = MidY
          Else
          ' If the midpoint is at or below the left, then midpoint is the new left. (We prefer the right half of the search region.)
            ' Get the new Left values...
            LeftX = MidX
            LeftY = MidY
          End If
        End If
        ' We just finished another iteration.
        IterationCount = IterationCount + 1
      Loop
    End If
  End If
  ' At this point ABC has been sorted and its values meet the derivative requirement.
  ' The calling function only cares about the search region, which is bounded by two of the three ABC points.
  ' If the middle points derivative is positive, we're working with the Left and Middle points.
  If ABC(2, 3) > 0 Then
    LR(1, 1) = ABC(1, 1)
    LR(1, 2) = ABC(1, 2)
    LR(1, 3) = ABC(1, 3)
    LR(2, 1) = ABC(2, 1)
    LR(2, 2) = ABC(2, 2)
    LR(2, 3) = ABC(2, 3)
  Else
    ' The derivative must have been negative or zero, so we're working with the Middle and Right points.
    LR(1, 1) = ABC(2, 1)
    LR(1, 2) = ABC(2, 2)
    LR(1, 3) = ABC(2, 3)
    LR(2, 1) = ABC(3, 1)
    LR(2, 2) = ABC(3, 2)
    LR(2, 3) = ABC(3, 3)
  End If
  ' Just return LR, which contains two points and their derivatives.
  ' The derivatives have opposite signs, as required for the successive quadratic approximation algorithm unless
  ' the constraint is active.  If the constraint is active, the return value is the point at the constraint.
  EnforceDerivativeRequirements = LR
End Function


' Returns the index of the minimum value in an array
Function FindMinValueIndex(A() As Double) As Long
  Dim i As Long
  Dim CurrentMin As Double
  
  ' Assume the first item has the minimum value...
  CurrentMin = A(LBound(A))
  FindMinValueIndex = LBound(A)
  ' Loop once for each remaining value.
  For i = LBound(A) + 1 To UBound(A)
    ' Is this one better?
    If A(i) < CurrentMin Then
      ' Remember its index and value.
      FindMinValueIndex = i
      CurrentMin = A(i)
    End If
  Next
End Function


' Inelegant way to check if a dynamic array is initialized.
Function IsInitialized(ByRef ArrayToCheck) As Boolean
  Dim i As Long
  
  ' Assume the array is not initialized.
  IsInitialized = False
  On Error GoTo ExitRoutine
  i = UBound(ArrayToCheck)
  If i > 0 Then
    ' If there is no error, it was initialized.
    IsInitialized = True
  End If
ExitRoutine:
  On Error GoTo 0
End Function


' Returns the maximum of two numbers.
Function Max(A, B) As Variant
  ' Assume A is higher.
  Max = A
  ' If B is higher, return that instead...
  If B > A Then
    Max = B
  End If
End Function


'won't navigate through precedents in closed workbooks
'won't navigate through precedents in protected worksheets
'won't identify precedents on hidden sheets
Public Function GetAllPrecedents(ByRef rngToCheck As Range) As Object
  Dim dicAllPrecedents As Object
  Dim strKey As String
  
  Set dicAllPrecedents = CreateObject("Scripting.Dictionary")
  ' Initiate the search for precedent cells.  The result is stored in dicAllPrecedents and transferred to the function return value.
  Call GetPrecedents(rngToCheck, dicAllPrecedents, 1)
  Set GetAllPrecedents = dicAllPrecedents
End Function


' This converts a range (of possibly more than one cell) to a series of calls to GetCellPrecedents.
Private Sub GetPrecedents(ByRef rngToCheck As Range, ByRef dicAllPrecedents As Object, ByVal lngLevel As Long)
  Dim rngCell As Range
  Dim rngFormulas As Range
  ' Don't check further if the cell's worksheet is protected.
  ' Note the misnamed property ProtectContents that should be named something like ContentsAreProtected.
  If Not rngToCheck.Worksheet.ProtectContents Then
    ' Is there more than one cell in this range?
    If rngToCheck.Cells.CountLarge > 1 Then
      On Error Resume Next
      ' Only check the cells that have formulas in them.
      Set rngFormulas = rngToCheck.SpecialCells(xlCellTypeFormulas)
      On Error GoTo 0
    Else
      ' This must have been only one cell (not a range of many cells).
      ' Does this cell contain a formula?
      If rngToCheck.HasFormula Then
        ' This has a formula, so we want to check for its precedents.
        Set rngFormulas = rngToCheck
      End If
    End If
    ' At this point rngFormulas either contains nothing or contains one or more cells with formulas.
    ' Did we find anything?
    If Not rngFormulas Is Nothing Then
      ' Iterate once for each cell with a formula.
      For Each rngCell In rngFormulas.Cells
        ' Start the whole process for this cell.  The MMM version arrives at the correct result.
        Call GetCellPrecedentsMMM(rngCell, dicAllPrecedents, lngLevel)
      Next rngCell
      ' We're done with these cells (though we may come across this worksheet again).
      rngFormulas.Worksheet.ClearArrows
    End If
  End If
End Sub


' Compiles a list of precedents to a single cell.  The result is stored in dicAllPrecedents.
Private Sub GetCellPrecedentsMMM(ByRef rngCell As Range, ByRef dicAllPrecedents As Object, ByVal lngLevel As Long)
  Dim lngArrow As Long
  Dim lngLink As Long
  Dim ContinueLookingForArrows As Boolean
  Dim strPrecedentAddress As String
  Dim rngPrecedentRange As Range
  
  ' The NavigateArrow method takes numerical "arrow" and "link" parameters.
  ' Loop through the arrows.  Then loop through the links.  Each valid arrow must have at least one valid link.
  ' The Excel object model doesn't provide a function that returns the number of valid arrows.  That would be too easy.
  ' It doesn't provide a function to return the number of valid links for a given arrow.  That also would be too easy.
  ' When you exceed the number of valid links for a given arrow, you get an error message. (And there may be more valid arrows to look through.)
  ' When you exceed the number of valid arrows, you get a reference back to the cell you are searching. (At which point there are no more valid arrows or links.)
  
  ' We haven't checked any arrows yet.
  lngArrow = 0
  ' Outer Do loop - Loop for arrows
  Do
    ' We're checking another arrow...
    lngArrow = lngArrow + 1
    ' Assume that we should not move to the next arrow after this loop.
    ' (This will change later if at least one valid result is found.)
    ContinueLookingForArrows = False
    ' ...but we haven't checked any links yet.
    lngLink = 0
    ' Inner Do loop - Loop for links
    Do
      ' We're checking the next link for this particular arrow.
      lngLink = lngLink + 1
      ' For some reason Excel skips some of the precedents if we don't do this again and again in each iteration.
      rngCell.ShowPrecedents
      ' Rather than tell us how far to search, Excel expects us to generate an error message and catch it when it arrives.
      On Error Resume Next
      ' Attempt to find a precedent with this arrow and link number.
      Set rngPrecedentRange = rngCell.NavigateArrow(True, lngArrow, lngLink)
      ' If this generated an error, the arrow/link combination must not refer to a valid precedent cell.
      ' In that case, stop looking for more links to go with this arrow.
      If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Do
      End If
      On Error GoTo 0
      ' If we got this far, there was no error, which means that the arrow/link combination produces a valid cell reference
      ' (but not necessarily a reference to a precedent).
      strPrecedentAddress = rngPrecedentRange.Address(False, False, xlA1, True)
      ' The only other thing that can go wrong is that the arrow/link combination references rngCell itself.
      ' This reveals an inconsistency in the NavigateArrow function's behavior.
      ' If one thing goes wrong, you get an error.  If something else goes wrong, you get a result that doesn't tell you anything.
      ' If this "precedent" references the cell we were searching, then we've exhausted the possible arrow values.
      If strPrecedentAddress = rngCell.Address(False, False, xlA1, True) Then
        ' This exits the inner Do loop.
        Exit Do
      Else
        ' The arrow/link combination produced a useful result, so there may be more valid arrows after this one.
        ' When this inner Do loop (for links) finishes, continue iterating in the outer Do loop (for more arrows).
        ContinueLookingForArrows = True
        ' If this is already in the list of precedents, its level (and its precedents' levels) may need to be updated.
        If dicAllPrecedents.Exists(strPrecedentAddress) Then
          ' Does the dictionary list a shallower level?  (If so, update it.  If not, leave it alone.)
          If dicAllPrecedents.Item(strPrecedentAddress) < lngLevel Then
            ' Replace the existing level with the updated, deeper level.
            dicAllPrecedents.Item(strPrecedentAddress) = lngLevel
            ' The precedent cell's own precedent cells also need to be updated.
            Call GetPrecedents(rngPrecedentRange, dicAllPrecedents, lngLevel + 1)
          End If
        Else
          ' This item must not be in the dictionary.
          ' Add this item and its precedents as usual.
          Call dicAllPrecedents.Add(strPrecedentAddress, lngLevel)
          Call GetPrecedents(rngPrecedentRange, dicAllPrecedents, lngLevel + 1)
        End If
      End If
    Loop
  ' If ContinueLookingForArrows is False, that marks the end of this branch of recursive calls.
  Loop While ContinueLookingForArrows
End Sub


' Recalculates an array of Range objects in order from first to last.
Sub RecalculateRanges(RangeArray() As Range)
  Dim i As Long
  For i = LBound(RangeArray) To UBound(RangeArray)
    Call RangeArray(i).Calculate
  Next
End Sub


' Returns an array of Range objects that are precedents to another Range.
' Each Range is a single cell.  Proper recalculation of the parent Range (which these were calculated from) proceeds from first to last.
Function ArrangePrecedents(PrecedentsDictionary As Object) As Range()
  Dim i As Long
  Dim j As Long
  Dim ItemsArray() As Variant
  Dim KeysArray() As Variant
  Dim MaxLevel As Long
  Dim ThisIndex As Long
  Dim Result() As Range
  
  ' Retrieve the dictionary data as arrays.
  ItemsArray = PrecedentsDictionary.Items()
  KeysArray = PrecedentsDictionary.Keys()
  ' Find the maximum level of any item in the array.
  MaxLevel = 0
  ' Loop once for each item and search for the highest value.  That will be MaxLevel.
  For i = LBound(ItemsArray) To UBound(ItemsArray)
    If ItemsArray(i) > MaxLevel Then
      MaxLevel = ItemsArray(i)
    End If
  Next
  ' The result must have the same size as the dictionary keys.
  ReDim Result(LBound(KeysArray) To UBound(KeysArray))
  ' We haven't chosen an index, but this will be incremented before it is first used.
  ' As we keep adding ranges to Result, we need to keep track of which element we used last.
  ThisIndex = LBound(Result) - 1
  ' Start with the deepest level MaxLevel and work down to 1.
  For i = MaxLevel To 1 Step -1
    ' Loop once for each item that we might add to the array.
    For j = LBound(ItemsArray) To UBound(ItemsArray)
      ' If this item has the level we're looking for, we need to add it to the result.
      If ItemsArray(j) = i Then
        ' Use the next location in Result and add the Range corresponding to the address string.
        ThisIndex = ThisIndex + 1
        Set Result(ThisIndex) = Application.Evaluate(KeysArray(j))
      End If
    Next
  Next
  ArrangePrecedents = Result
End Function
