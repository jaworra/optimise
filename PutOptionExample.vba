' Calculate the put derivative baseline value v.
Public Function GetPutBaseline(S As Double, K As Double, T As Double, r As Double, Delta As Double, Sigma As Double) As Double
  Dim d1 As Double
  Dim d2 As Double
  ' Calculate parameters.
  d1 = Log(S / K) + (r - Delta + (Sigma ^ 2) / 2) * T
  d1 = d1 / (Sigma * T ^ 0.5)
  d2 = d1 - Sigma * T ^ 0.5
  ' Calculate P.
  GetPutBaseline = -S * Exp(-Delta * T) * Application.WorksheetFunction.NormSDist(-d1)
  GetPutBaseline = GetPutBaseline + K * Exp(-r * T) * Application.WorksheetFunction.NormSDist(-d2)
End Function

