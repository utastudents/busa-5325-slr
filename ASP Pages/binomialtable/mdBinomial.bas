Attribute VB_Name = "mdBinomial"
Option Explicit
Public Function CumBinProb(ln As Long, lx As Long, dP As Double) As Double
   Dim intIndex As Long
   Dim lngX As Long
   Dim dblP As Double
   
   Dim dblBinProb As Double
   CumBinProb = 0
   For intIndex = 0 To lx
     lngX = intIndex
     dblP = dP
     dblBinProb = BinomialProb2(ln, lngX, dblP)
     CumBinProb = CumBinProb + dblBinProb
     Debug.Print intIndex, dblBinProb, CumBinProb
   Next intIndex
End Function

Public Function BinomialProb2(ln As Long, lx As Long, dP As Double) As Double
   Dim iloopcounter As Long
   
   BinomialProb2 = 1#
   
   'check for incorrect input
   If lx > ln Then
     MsgBox "Count is larger than sample size"
     Exit Function
   End If
   
   'calculate N-X
   Dim lNminusX As Long
   lNminusX = ln - lx
   'switch values if x>lNminusX
   If lx > lNminusX Then
     lx = lNminusX
     lNminusX = ln - lx
     dP = 1# - dP
   End If
   
   'Computer the product of integers lMin to lMax
   BinomialProb2 = 1#
   For iloopcounter = 1 To lNminusX
      If iloopcounter <= lx Then
          BinomialProb2 = BinomialProb2 * dP * (1# - dP) * (ln + 1# - iloopcounter) / (lx + 1# - iloopcounter)
      Else
          BinomialProb2 = BinomialProb2 * (1# - dP)
      End If
   Next iloopcounter
End Function
Public Function LogBinomial(ln As Long, lx As Long, dP As Double) As Double
   Dim iloopcounter As Long
   
   LogBinomial = 1#
   
   'check for incorrect input
   If lx > ln Then
     MsgBox "Count is larger than sample size"
     Exit Function
   End If
   
   'calculate N-X
   Dim lNminusX As Long
   lNminusX = ln - lx
   
   'Computer the product of integers lMin to lMax
   LogBinomial = 1#
   Dim dNumberOfCombinations As Double
   
     dNumberOfCombinations = logFactorial(ln) - logFactorial(lx) - logFactorial(ln - lx)
     LogBinomial = dNumberOfCombinations + lx * Log(dP) + (ln - lx) * Log(1# - dP)
   
   End Function

Public Function logFactorial(iAmount As Long) As Double
   Dim iloopcounter As Long
   logFactorial = 0#
   If iAmount <> 0 Then
     For iloopcounter = 1 To iAmount
         logFactorial = logFactorial + Log(iloopcounter)
     Next iloopcounter
   Else
      logFactorial = 0#
   End If
End Function
