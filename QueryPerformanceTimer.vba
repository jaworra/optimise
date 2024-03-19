' Timer provided as-is at:
' http://support.microsoft.com/kb/172338

Option Base 1

Option Explicit

' Yes.  We have to convert to a Boolean type explicitly or VBA freaks out and doesn't evaluate the expression correctly.
#If CBool(VBA7) Then
  ' The compiler is looking for PtrSafe in front of these functions, though their signatures are the same.
  Declare PtrSafe Function QueryPerformanceCounter Lib "Kernel32" (x As Currency) As Boolean
  Declare PtrSafe Function QueryPerformanceFrequency Lib "Kernel32" (x As Currency) As Boolean
  Declare PtrSafe Function GetTickCount Lib "Kernel32" () As Long
  Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
  ' The compiler is not looking for PtrSafe in front of these functions.
  Declare Function QueryPerformanceCounter Lib "Kernel32" (x As Currency) As Boolean
  Declare Function QueryPerformanceFrequency Lib "Kernel32" (x As Currency) As Boolean
  Declare Function GetTickCount Lib "Kernel32" () As Long
  Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

Sub Test_Timers()
Dim Ctr1 As Currency, Ctr2 As Currency, Freq As Currency
Dim Count1 As Long, Count2 As Long, Loops As Long
'
' Time QueryPerformanceCounter
'
  If QueryPerformanceCounter(Ctr1) Then
    QueryPerformanceCounter Ctr2
    Debug.Print "Start Value: "; Format$(Ctr1, "0.0000")
    Debug.Print "End Value: "; Format$(Ctr2, "0.0000")
    QueryPerformanceFrequency Freq
    Debug.Print "QueryPerformanceCounter minimum resolution: 1/" & Freq * 10000; " sec"
    Debug.Print "API Overhead: "; (Ctr2 - Ctr1) / Freq; "seconds"
  Else
    Debug.Print "High-resolution counter not supported."
  End If
'
' Time GetTickCount
'
  Debug.Print
  Loops = 0
  Count1 = GetTickCount()
  Do
    Count2 = GetTickCount()
    Loops = Loops + 1
  Loop Until Count1 <> Count2
  Debug.Print "GetTickCount minimum resolution: " & (Count2 - Count1); "ms"
  Debug.Print "Took" & Loops & "loops"
'
' Time timeGetTime
'
  Debug.Print
  Loops = 0
  Count1 = timeGetTime()
  Do
    Count2 = timeGetTime()
    Loops = Loops + 1
  Loop Until Count1 <> Count2
  Debug.Print "timeGetTime minimum resolution: " & (Count2 - Count1); "ms"
  Debug.Print "Took" & Loops&; "loops"
End Sub


Private Sub Time_Addition()
  Dim Ctr1 As Currency, Ctr2 As Currency, Freq As Currency
  Dim Overhead As Currency, A As Long, i As Long
  QueryPerformanceFrequency Freq
  QueryPerformanceCounter Ctr1
  QueryPerformanceCounter Ctr2
  Overhead = Ctr2 - Ctr1        ' determine API overhead
  QueryPerformanceCounter Ctr1  ' time loop
  For i = 1 To 100
    A = A + i
  Next i
  QueryPerformanceCounter Ctr2
  Debug.Print "("; Ctr1; "-"; Ctr2; "-"; Overhead; ") /"; Freq
  Debug.Print "100 additions took";
  Debug.Print (Ctr2 - Ctr1 - Overhead) / Freq; "seconds"
End Sub


Public Function DisplayElapsedTime(ActivityName As String, Time1 As Currency, Time2 As Currency)
  Dim Overhead As Currency
  Dim Frequency As Currency
  Dim Ctr1 As Currency
  Dim Ctr2 As Currency
  Dim ElapsedTime As Currency
  Call QueryPerformanceCounter(Ctr1)
  Call QueryPerformanceCounter(Ctr2)
  Overhead = Ctr2 - Ctr1
  Call QueryPerformanceFrequency(Frequency)
  ElapsedTime = Time2 - Time1 - Overhead
  DisplayElapsedTime = ActivityName & ": " & Round(ElapsedTime / Frequency, 3)
End Function


Sub TestTimer()
  Dim Ctr1 As Currency
  Dim Ctr2 As Currency
  Dim i
  Dim j
  Call QueryPerformanceCounter(Ctr1)
  ' Do something slow.
  For i = 1 To 10000000
    j = j + 1
  Next
  Call QueryPerformanceCounter(Ctr2)
  ' Show the user how long it took.
  Debug.Print DisplayElapsedTime("Test", Ctr1, Ctr2)
End Sub


Sub TestArrayWriteTime()
  Dim Time1 As Currency
  Dim Time2 As Currency
  Dim i As Long
  Dim j As Long
  Const Length As Long = 5000
  Dim MyArray(1 To Length, 1 To Length) As Double
  Call QueryPerformanceCounter(Time1)
  For i = 1 To 5000
    For j = 1 To 5000
      MyArray(j, i) = MyArray(i, j) + 9
    Next
  Next
  Call QueryPerformanceCounter(Time2)
  Debug.Print DisplayElapsedTime("Writing 25E6 Points", Time1, Time2)
End Sub

