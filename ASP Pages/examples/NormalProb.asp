
<HTML>
<HEAD><TITLE>Normal Probability</TITLE>


</HEAD>
<BODY>
<form action=ProbOfXSol.asp method=POST enctype="application/x-www-form-urlencoded">
<%
   Dim Question()

    ReDim Question(4)
    NumAttempts = 0
    session.timeout =30
   'Creation of scenarios
       variable = Array("age", "mpg ", "cost", "sales", "price", "height")
        measure = Array("years", "miles", "dollars", "dollars", "dollars", "inches")
        observation = Array("buyers", "cars", "trips", "LCD panels", "tickets", "students")
        mean = Array(41, 27, 115, 632, 23, 35)
        st_dev = Array(7, 5, 15, 66, 6, 1)
     
     Randomize
             Scenario = Round(Rnd() * 6 + 0.5, 0) - 1
             
    'This scenarios values
        x = variable(Scenario)
        mea = measure(Scenario)
        obsr = observation(Scenario)
        mu = mean(Scenario)
        sig = st_dev(Scenario)

       
    'Determing a random < range and its prob
        z = CDbl(Round((Rnd() - 0.5) * 6, 2))
        LessThanValue = mu + z * sig
        LessThanProb = alnorm(z)
        Question(1) = "1. What is the probability of finding " & obsr & " with a value of " & x & _
                  " less than " & LessThanValue & " " & mea & "?"
      
      'Determing a random > range and its prob
        z = Round((Rnd() - 0.5) * 6, 2)
        GrThanValue = mu + z * sig
        GrThanProb = 1 - alnorm(z)
        Question(2) = "2. What is the probability of finding " & obsr & " with a value of " & x & _
                   " greater than " & GrThanValue & " " & mea & "?"
           
    'Determining a random range and then finding its prob
        zupper = RandBetween(-2.8, 3)
        zlower = RandBetween(-3, zupper)
        uppervalue = mu + zupper * sig
        lowervalue = mu + zlower * sig
        BetweenProb = alnorm(zupper) - alnorm(zlower)
        Question(3) = "3. What is the probability of finding " & obsr & " with a value of " & x & _
            " between " & lowervalue & " and " & uppervalue & " " & mea & "?"

    'Determining a determining a percentile and then finding its value of x
        choice = Round(Rnd() * 7 + 0.5, 0) - 1
        Percentile = Round(Rnd() * 0.9,2) + 0.05
        z_for_pctile = NormSInv(Percentile)
        X_for_pctile = mu + z_for_pctile * sig
        EndingDecimal=round((Percentile-ROUND(INT(Percentile*10)/10,1))*100,0)
        Select case endingdecimal
           case 1
             postfix = "st"
           case 2
             postfix = "nd"
           case 3 
             postfix = "rd"
           case else
             postfix = "th"
         end select
        if ((percentile>=0.11) and (percentile<=0.13)) then postfix = "th"

        Question(4) = "4. What value of " & x & " falls on the " & Percentile*100 & postfix & " percentile?"
  
     paragraph = "Suppose the " & x & " of " & obsr & " were collected and the data follows " & _
     "the bell curve (a normal distribution). The " & obsr & " had an average " & x & _
     " of " & mu & " " & mea & " and a standard deviation of " & sig & _
     ". "

        Session("x")=x 
        Session("mea")=mea 
        Session("obsr")=obsr 
        Session("mu")=mu 
        Session("sig")=sig 
        Session("bounds")=bounds 
        Session("Prob")=Prob 
        Session("prless")=prless 
        Session("prGreater")=prGreater 
        Session("LessThanProb")=LessThanProb 
        Session("GrThanProb")=GrThanProb 
        Session("LessThanValue")=LessThanValue 
        Session("GrThanValue")=GrThanValue 
        Session("UpperValue")=UpperValue 
        Session("LowerValue")=LowerValue 
        
        Session("BetweenProb")=BetweenProb
        Session("Percentile") = Percentile
        Session("Postfix") = PostFix
        Session("X_for_pctile")=X_for_pctile 
        Session("paragraph")=paragraph
        Session("NumAttempts") = numattempts

        Name=""
        if Session("Name")<>"" then name=replace(session("Name")," ","-")


    Function alnorm(x) 
'
'         Algorithm AS66 Applied Statistics (1973) vol22 no.3
'
'       Evaluates the tail area of the standardised normal curve
'       from x to infinity if upper is .true. or
'       from minus infinity to x if upper is .false.
'
     up=FALSE
     
     '*** machine dependent ants
      

      zero = 0.
      one = 1.
      half = 0.5
      ltone = 7.
      utzero = 18.66
      con = 1.28
      p = 0.398942280444
      q = 0.39990348504
      r = 0.398942280385
      a1 = 5.75885480458
      a2 = 2.62433121679
      a3 = 5.92885724438
      b1 = -29.8213557807
      b2 = 48.6959930692
      c1 = -0.000000038052
      c2 = 0.000398064794
      c3 = -0.151679116635
      c4 = 4.8385912808
      c5 = 0.742380924027
      c6 = 3.99019417011
      d1 = 1.00000615302
      d2 = 1.98615381364
      d3 = 5.29330324926
      d4 = -15.1508972451
      d5 = 30.789933034
'
      up = upper
      z = x


      If (z < zero) Then
         up = Not (up)
         z = -z
      End If
  
      If ((z < ltone Or up) And z < utzero) Then
         y = half * z * z
         If (z > con) Then
            part1 = d4 / (z + c5 + d5 / (z + c6))
            part2 = d3 / (z + c4 + part1)
            part3 = d2 / (z + c3 + part2)
            alnorm = r * Exp(-y) / (z + c1 + (d1 / (z + c2 + part3)))
         Else
            part1 = b2 / (y + a3)
            part2 = b1 / (y + a2 + part1)
            alnorm = half - z * (p - q * (y / (y + a1 + part2)))
         End If
      Else
        alnorm = zero
      End If
      
      If up = False Then alnorm = one - alnorm
      
     

      End Function

   Function NormSInv(p) 

    a1 = -39.6968302866538: a2 = 220.946098424521: a3 = -275.928510446969
    a4 = 138.357751867269: a5 = -30.6647980661472: a6 = 2.50662827745924
    b1 = -54.4760987982241: b2 = 161.585836858041: b3 = -155.698979859887
    b4 = 66.8013118877197: b5 = -13.2806815528857: c1 = -7.78489400243029E-03
    c2 = -0.322396458041136: c3 = -2.40075827716184: c4 = -2.54973253934373
    c5 = 4.37466414146497: c6 = 2.93816398269878: d1 = 7.78469570904146E-03
    d2 = 0.32246712907004: d3 = 2.445134137143: d4 = 3.75440866190742
    p_low = 0.02425: p_high = 1 - p_low

    If p < 0 Or p > 1 Then
      'Err.Raise vbObjectError, , "NormSInv: Argument out of range."
    ElseIf p < p_low Then
      q = Sqr(-2 * Log(p))
      NormSInv = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
         ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
    ElseIf p <= p_high Then
      q = p - 0.5: r = q * q
      NormSInv = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / _
         (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)
    Else
      q = Sqr(-2 * Log(1 - p))
      NormSInv = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
         ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
      End If

  
   End Function

 Function RandBetween(xl, xu)
      RandBetween = Round(Rnd() * (xu - xl), 2) + xl
    End Function
     

  sub PrHtml(strHtml)
     response.write(strhtml)
  end sub
%>
  
 <h1 align="center"> Exercise on the Probability of a Normally Distributed Random Variable</h1>
   
   
  <table align="center" width="600">
    <tr>
     <td>
      Enter your name, last name first
		<input style="width:350px" name="name" value=<%=Name%>>
      	        
     </td>
    </tr>
   <tr>
   </table><br>

 <table align="center" border="1" width="600">
     <tr>
       <td align="left" style="word-wrap: break-word"> 
         <%response.write(paragraph)%>  <br>
       </td>
     </tr>
     </table>	
     <br>

 <table align="center" border="1"  width="600"> 
   <tr><td>
      <b> Directions: Fill in the blanks and click the button to have your answers graded.</b><br>
      A. Do not round intermediate calculations. <br>
      B. For questions 1-3 enter a number between 0 and 1 rounding it to four places.<br>
      C. Round your answer to question 4 to two decimal places.<br>
      D. You have <% response.write(session.timeout) %> minutes to complete your work before session
      timeout. After that your work <br> will not be graded correctly. </p>
   </td></tr></table>

 <table align="center" >  
   
   <%
      for i=1 to 4
         PrHtml("<tr><td>" & Question(i) & " = <input  name=" & chr(34) & "AnsQues" & i & chr(34) & "></td></tr>" )
       next
    %>
    </table><br>     

<p align="center">
  <input type="submit" value="Please click here to check answers" name="END">
</p>

<p align="center">
   Click <a align="center" href="<%= request.servervariables("script_name") %>">here</a> for a new exercise
</p>

 


</BODY>
</HTML>

