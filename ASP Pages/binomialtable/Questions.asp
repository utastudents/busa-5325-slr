<html>
<body>
 <%strGraphType = Request.form("GO")
   dim strHtmlPage()
   Randomize
   'Input, check and look up zscore 1
   intX1 = int(request.form("GivenX1"))
   intX2 = int(request.form("GivenX2"))
   intN = int(request.form("GivenN"))
   dblP = request.form("GivenP")
  
   'Check for impossible values
    
   if dblP > 1 then 
     response.write "ERROR: Given value of P was too large and had to be changed!<p>"  
     dblP = 1.
   end if

   if dblP < 0 then 
     response.write "ERROR: Given value of P was too small and had to be changed!<p>"  
     dblP = 0.   
   end if
   
    
    redim strHtmlPage(4)
    strHtmlPage(0) = strHeading

  select case strgraphtype
    case "A"
      'probability that X is less than or equal to a value X1
      strHtmlPage(0) = "<h2> Find Pr ( X &lt= x1) </h2> Change the values in the text boxes and click calculate" 
       'check for impossible values
       if intX1 > intN then 
           response.write "ERROR: Given value of X1 was too large and had to be changed! <p>"
           intX1=intN
       end if
       if intX1 < 0 then 
           response.write "ERROR: Given value of X1 was too small and had to be changed! <p>"
           intX1=0
       end if
      BinProb = round(cumbinprob(intn, intx1, dblP),3)
       'Fill in message
       strHtmlPage(1) = "Determine the probability that X is less than or equal to  " 
       strHtmlPage(2) =  "<INPUT NAME='givenX1' VALUE='" & intX1 & "' size=4> <p>"
       strAnswer = "Pr ( X <= " & intX1 & ") = " & BinProb 
    case "B"
    'probability that X is less than a value X1
    strHtmlPage(0) = "<h2> Find Pr ( X &lt x1) </h2> Change the values in the text boxes and click calculate" 
    
    'check for impossible values
       if intX1 > intN then 
           response.write "ERROR: Given value of X1 was too large and had to be changed! <p>"
           intX1=intN
       end if
       if intX1 < 1 then 
           response.write "ERROR: Given value of X1 was too small and had to be changed! <p>"
           intX1=1
       end if
     BinProb = round(cumbinprob(intn, intx1-1, dblP),3)
      'fill in message
       strHtmlPage(1) = "Determine the probability that X is less than " 
       strHtmlPage(2) =  "<INPUT NAME='givenX1' VALUE='" & intX1 & "' size=4> <p>"
       strAnswer = "Pr ( X < " & intX1 & ") = Pr ( X <= " & intX1-1 & " ) = " & BinProb 
   case "C"
    'probability that X is greater than a value X1
    strHtmlPage(0) = "<h2> Find Pr ( X &gt x1) </h2> Change the values in the text boxes and click calculate" 
    
    'check for impossible values
       if intX1 > intN then 
           response.write "ERROR: Given value of X1 was too large and had to be changed! <p>"
           intX1=intN
       end if
       if intX1 < 0 then 
           response.write "ERROR: Given value of X1 was too small and had to be changed! <p>"
           intX1=0
       end if
     BinProb = 1- round(cumbinprob(intn, intx1, dblP),3)
      'fill in message
       strHtmlPage(1) = "Determine the probability that X is greater than " 
       strHtmlPage(2) =  "<INPUT NAME='givenX1' VALUE='" & intX1 & "' size=4> <p>"
       strAnswer = "Pr ( X > " & intX1 & ") = 1 - Pr ( X <= " & intX1 & " ) = " & binprob 
  case "D"
    'probability that X is greater than or equal to a value X1
    strHtmlPage(0) = "<h2> Find Pr ( X &gt= x1) </h2> Change the values in the text boxes and click calculate" 
    
    'check for impossible values
       if intX1 > intN then 
           response.write "ERROR: Given value of X1 was too large and had to be changed! <p>"
           intX1=intN
       end if
       if intX1 < 0 then 
           response.write "ERROR: Given value of X1 was too small and had to be changed! <p>"
           intX1=0
       end if
         strHtmlPage(1) = "Determine the probability that X is greater than or equal to " 
         strHtmlPage(2) =  "<INPUT NAME='givenX1' VALUE='" & intX1 & "' size=4> <p>"
   
      if intX1>0 then
        BinProb = 1- round(cumbinprob(intn, intx1-1, dblP),3)
        'fill in message
         strAnswer = "Pr ( X >= " & intX1 & ") = 1 - Pr ( X <= " & intX1-1 & " ) = " & binprob 
      else
        'fill in message
         strAnswer = "Pr ( X >= " & intX1 & ") = 1"
      end if  
  case "E"
    'probability that X is equal to a value X1
    strHtmlPage(0) = "<h2> Find Pr ( X = x1) </h2> Change the values in the text boxes and click calculate" 
    
    'check for impossible values
       if intX1 > intN then 
           response.write "ERROR: Given value of X1 was too large and had to be changed! <p>"
           intX1=intN
       end if
       if intX1 < 0 then 
           response.write "ERROR: Given value of X1 was too small and had to be changed! <p>"
           intX1=0
       end if
         strHtmlPage(1) = "Determine the probability that X is equal to " 
         strHtmlPage(2) =  "<INPUT NAME='givenX1' VALUE='" & intX1 & "' size=4> <p>"
   
       BinProb1 = round(cumbinprob(intn, intx1, dblP),3)

      if intX1>0 then
        BinProb2 = round(cumbinprob(intn, intx1-1, dblP),3)
        BinProb = round(binprob1-binprob2,3)
        'fill in message
         strAnswer = "Pr ( X = " & intX1 & ") = Pr ( X <= " & intX1 & " )" & _
                    " - Pr ( X <= " & intX1-1 & " ) = " & _
                    binprob1 & " - " & binprob2 & " = " & binprob
      else
        'fill in message
         strAnswer = "Pr ( X = " & intX1 & ") = " & binprob1
      end if  
 case "F"
    'probability that X is greater than or equal to a value X1 and less than or equal to X2
    strHtmlPage(0) = "<h2> Find Pr (  x1 &lt = X &lt= x2) </h2> Change the values in the text boxes and click calculate" 
    
    'check for impossible values
       if intX1 > intN then 
           response.write "ERROR: Given value of X1 was too large and had to be changed! <p>"
           intX1=intN
       end if
       if intX1 < 0 then 
           response.write "ERROR: Given value of X1 was too small and had to be changed! <p>"
           intX1=0
       end if
       if intX2 > intN then
           response.write "ERROR: Given value of X2 was too large and had to be changed! <p>"
           intX2=intN
       end if
       if intX2 < 0 then 
           response.write "ERROR: Given value of X2 was too small and had to be changed! <p>"
           intX2=0
       end if
       if intX1 > intX2 then 
           response.write "ERROR: Given value of x1 was larger than x2 and had to be changed! <p>"
           intTemp = intX1            
           intX1=intX2
           intX2 = intTemp
       end if


         strHtmlPage(1) = "Determine the probability that X is greater than or equal to " & _
                           "<INPUT NAME='givenX1' VALUE='" & intX1 & "' size=4> <p> "
         strHtmlPage(2) =  "and less than or equal to " & _
                           "<INPUT NAME='givenX2' VALUE='" & intX2 & "' size=4> <p> "
   
        BinProbX2 = round(cumbinprob(intn, intx2, dblP),3) 
      
      if intX1>0 then
        BinProbX1 =  round(cumbinprob(intn, intx1-1, dblP),3)
        BinProb = round(BinProbX2 - BinProbX1,3)
        'fill in message
         strAnswer = "Pr ( " & intX1 & " <= X <= " & intX2 & ") = " & _
                     "Pr ( X <= " & intX2 & " ) - Pr ( X <= " & intX1-1 & " ) = " & _
                     binprobx2 & " - " & binprobx1 & " = " & binprob 
      else
        'fill in message
         strAnswer = "Pr ( X <= " & intX2 & ") = " & binprobx2
      end if  
case "G"
    'probability that X is greater than or equal to a value X1 and less than X2
    strHtmlPage(0) = "<h2> Find Pr (  x1 &lt= X &lt x2) </h2> Change the values in the text boxes and click calculate" 
    
    'check for impossible values
       if intX1 >= intN then 
           response.write "ERROR: Given value of X1 was too large and had to be changed! <p>"
           intX1=intN-1
       end if
       if intX1 < 0 then 
           response.write "ERROR: Given value of X1 was too small and had to be changed! <p>"
           intX1=0
       end if
       if intX2 > intN then
           response.write "ERROR: Given value of X2 was too large and had to be changed! <p>"
           intX2=intN
       end if
       if intX2 < 1 then 
           response.write "ERROR: Given value of X2 was too small and had to be changed! <p>"
           intX2=1
       end if
       if intX1 > intX2-1 then 
           response.write "ERROR: Given value of x1 was larger than x2-1 and had to be changed! <p>"
           intTemp = intX1            
           intX1=intX2-1
           intX2 = intTemp
       end if


         strHtmlPage(1) = "Determine the probability that X is greater than or equal to " & _
                           "<INPUT NAME='givenX1' VALUE='" & intX1 & "' size=4> <p> "
         strHtmlPage(2) =  "and less than " & _
                           "<INPUT NAME='givenX2' VALUE='" & intX2 & "' size=4> <p> "
   
        BinProbX2 = round(cumbinprob(intn, intx2-1, dblP),3) 
      
      if intX1>0 then
        BinProbX1 =  round(cumbinprob(intn, intx1-1, dblP),3)
        BinProb = round(BinProbX2 - BinProbX1,3)
        'fill in message
         strAnswer = "Pr ( " & intX1 & " <= X < " & intX2 & ") = " & _
                     "Pr ( X <= " & intX2-1 & " ) - Pr ( X <= " & intX1-1 & " ) = " & _
                     binprobx2 & " - " & binprobx1 & " = " & binprob 
      else
        'fill in message
         strAnswer = "Pr ( X <= " & intX2-1 & ") = " & binprobx2
      end if  
case "H"
    'probability that X is greater than X1 and less than or equal to X2
    strHtmlPage(0) = "<h2> Find Pr (  x1 &lt X &lt= x2) </h2> Change the values in the text boxes and click calculate" 
    
    'check for impossible values
       if intX1 >= intN-1 then 
           response.write "ERROR: Given value of X1 was too large and had to be changed! <p>"
           intX1=intN-1
       end if
       if intX1 < 0 then 
           response.write "ERROR: Given value of X1 was too small and had to be changed! <p>"
           intX1=0
       end if
       if intX2 > intN then
           response.write "ERROR: Given value of X2 was too large and had to be changed! <p>"
           intX2=intN
       end if
       if intX2 < 1 then 
           response.write "ERROR: Given value of X2 was too small and had to be changed! <p>"
           intX2=1
       end if
       if intX1 > intX2-1 then 
           response.write "ERROR: Given value of x1 was larger than x2-1 and had to be changed! <p>"
           intTemp = intX1            
           intX1=intX2-1
           intX2 = intTemp
       end if


         strHtmlPage(1) = "Determine the probability that X is greater than " & _
                           "<INPUT NAME='givenX1' VALUE='" & intX1 & "' size=4> <p> "
         strHtmlPage(2) =  "and less than or equal to " & _
                           "<INPUT NAME='givenX2' VALUE='" & intX2 & "' size=4> <p> "
   
        BinProbX2 = round(cumbinprob(intn, intx2, dblP),3) 
      
      if intX1>0 then
        BinProbX1 =  round(cumbinprob(intn, intx1, dblP),3)
        BinProb = round(BinProbX2 - BinProbX1,3)
        'fill in message
         strAnswer = "Pr ( " & intX1 & " < X <= " & intX2 & ") = " & _
                     "Pr ( X <= " & intX2 & " ) - Pr ( X <= " & intX1 & " ) = " & _
                     binprobx2 & " - " & binprobx1 & " = " & binprob 
      else
        'fill in message
         strAnswer = "Pr ( X <= " & intX2 & ") = " & binprobx2
      end if  
case "I"
    'probability that X is greater than X1 and less than X2
    strHtmlPage(0) = "<h2> Find Pr (  x1 &lt X &lt x2) </h2> Change the values in the text boxes and click calculate" 
    
    'check for impossible values
       if intX1 > intX2-2 then 
           response.write "ERROR: Given value of x1 was larger than x2-1 and had to be changed! <p>"
           intTemp = intX1            
           intX1=intX2-2
           intX2 = intTemp
       end if
       if intX1 >= intN-1 then 
           response.write "ERROR: Given value of X1 was too large and had to be changed! <p>"
           intX1=intN-2
       end if
       if intX1 < 0 then 
           response.write "ERROR: Given value of X1 was too small and had to be changed! <p>"
           intX1=0
       end if
       if intX2 > intN then
           response.write "ERROR: Given value of X2 was too large and had to be changed! <p>"
           intX2=intN
       end if
       if intX2 < 2 then 
           response.write "ERROR: Given value of X2 was too small and had to be changed! <p>"
           intX2=2
       end if
       

         strHtmlPage(1) = "Determine the probability that X is greater than " & _
                           "<INPUT NAME='givenX1' VALUE='" & intX1 & "' size=4> <p> "
         strHtmlPage(2) =  "and less than " & _
                           "<INPUT NAME='givenX2' VALUE='" & intX2 & "' size=4> <p> "
   
        BinProbX2 = round(cumbinprob(intn, intx2-1, dblP),3) 
      
      if intX1>1 then
        BinProbX1 =  round(cumbinprob(intn, intx1-1, dblP),3)
        BinProb = round(BinProbX2 - BinProbX1,3)
        'fill in message
         strAnswer = "Pr ( " & intX1 & " < X < " & intX2 & ") = " & _
                     "Pr ( X <= " & intX2-1 & " ) - Pr ( X <= " & intX1-1 & " ) = " & _
                     binprobx2 & " - " & binprobx1 & " = " & binprob 
      else
        'fill in message
         strAnswer = "Pr ( X <= " & intX2-1 & ") = " & binprobx2
      end if  


   end select 
    
      
    strHtmlPage(3) = " when N = <INPUT NAME='givenN' VALUE='" & intN & "' size=4> and "
    strhtmlPage(4) = "P = <INPUT NAME='givenP' VALUE='" & dblP & "' size=4> <p> "
   
%>
  <%=strHtmlPage(0) %>

 <Form action="cumbin.asp" method="post">

 <% for intIndex = 1 to 4
    response.write strHtmlPage(intIndex)
   next 
  %>
 
 <input Name="go" value="<%=strgraphtype%>" Type = "Hidden">          
 <INPUT TYPE=submit VALUE="Calculate">
 <p>
 
 <% 
     response.write strAnswer
   %>   
   
     
 </form>
<p>
<p>


<%

Public Function CumBinProb(ln , lx , dP ) 
   Dim intIndex 
   Dim lngX 
   Dim dblP 
   
   Dim dblBinProb 
   CumBinProb = 0
   For intIndex = 0 To lx
     lngX = intIndex
     dblP = dP
     dblBinProb = BinomialProb2(ln, lngX, dblP)
     CumBinProb = CumBinProb + dblBinProb
     'Debug.Print intIndex, dblBinProb, CumBinProb
   Next 
End Function

Public Function BinomialProb2(ln , lx , dP ) 
   Dim iloopcounter 
   
   BinomialProb2 = 1.
   
   'check for incorrect input
   If lx > ln Then
     'MsgBox "Count is larger than sample size"
     Exit Function
   End If
   
   'calculate N-X
   Dim lNminusX 
   lNminusX = ln - lx
   'switch values if x>lNminusX
   If lx > lNminusX Then
     lx = lNminusX
     lNminusX = ln - lx
     dP = 1. - dP
   End If
   
   'Computer the product of integers lMin to lMax
   BinomialProb2 = 1.
   For iloopcounter = 1 To lNminusX
      If iloopcounter <= lx Then
          BinomialProb2 = BinomialProb2 * dP * (1. - dP) * (ln + 1. - iloopcounter) / (lx + 1. - iloopcounter)
      Else
          BinomialProb2 = BinomialProb2 * (1. - dP)
      End If
   Next 
End Function
Public Function LogBinomial(ln , lx , dP ) 
   Dim iloopcounter 
   
   LogBinomial = 1.
   
   'check for incorrect input
   If lx > ln Then
     'MsgBox "Count is larger than sample size"
     Exit Function
   End If
   
   'calculate N-X
   Dim lNminusX 
   lNminusX = ln - lx
   
   'Computer the product of integers lMin to lMax
   LogBinomial = 1.
   Dim dNumberOfCombinations 
   
     dNumberOfCombinations = logFactorial(ln) - logFactorial(lx) - logFactorial(ln - lx)
     LogBinomial = dNumberOfCombinations + lx * Log(dP) + (ln - lx) * Log(1. - dP)
   
   End Function

Public Function logFactorial(iAmount ) 
   Dim iloopcounter 
   logFactorial = 0.
   If iAmount <> 0 Then
     For iloopcounter = 1 To iAmount
         logFactorial = logFactorial + Log(iloopcounter)
     Next 
   Else
      logFactorial = 0.
   End If
End Function
  
%>   
</body>
<a href="p2.asp">Return to Main Menu</a>
</html>
