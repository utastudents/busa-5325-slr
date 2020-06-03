
<HTML>
<HEAD><TITLE>Z CI on Pi</TITLE>



<HTML>
<HEAD><TITLE>Z CI on Pi</TITLE>


</HEAD>
<BODY>
<form action=ZCIforPiSol.asp method=POST enctype="application/x-www-form-urlencoded">
<%
   Dim Question()

    ReDim Question(4)
    NumAttempts = 0
    NumAttCon = 0
    session.timeout=30
   'Creation of scenarios
       variable = Array("young", "foreign ", "highway", "flat", "football", "graduate")
        measure = Array("years", "miles", "dollars", "dollars", "dollars", "inches")
        observation = Array("buyers", "cars", "trips", "LCD panels", "tickets", "students")
       CL_Levels = Array(0.80,0.90,0.95,0.98,0.99)
        Z_Values = Array(1.28, 1.645, 1.96, 2.33, 2.576)
   Randomize 
        Pi=round(randbetween(.2,.8),2)
        sig = round(sqr(Pi*(1.-Pi)),8)
     
        Scenario = Round(Rnd() * 6 + 0.5, 0) - 1
        z1 = CDbl(Round((Rnd() - 0.5) * 6, 2))
  

    'This scenarios values
        x = variable(Scenario)
        mea = measure(Scenario)
        obsr = observation(Scenario)
  

    'Determine a sample size between 20 and 100
        n = Round(randbetween(20, 50), 0)
       
     'Determine the Standard error value
        SE = round(sig/sqr(n),8)
        P = round(Pi + z1 * SE,2)  
    
    'Determining a determining middle proportion and then finding the two P values 
        SamSe = sqr(p*(1.-p)/n)
 
       ChosenProbIndex = round(rnd()*5+.5,0)-1
       CL = Cl_Levels(chosenProbIndex)

        'Find area less than the upper range
        UpperPercentile = 1.-(1-Cl)/2   
        Z = z_Values(chosenProbIndex)
        'Find area less than the upper range
        UpperPercentile = 1.-(1-Cl)/2   

        'P=abs(round(Pi+round(normsinv(rnd()*.9+.05),2)*se,2))
        MOE = round(z*samse,8)

        LowerB = round(P - z * samSE,2)
        UpperB = round(P + z * samSE,2)
        'prhtml("<p> moe = " & moe & " </p>")
        'prhtml("<p> Lower Bound = " & LowerB & ", " & "Upper Bound = " & UpperB & "</p>")

       

 
     paragraph = "Estimate the proportion of all " & obsr & " that are " & x & " with " & CL*100 & "% " & _
       "confidence. After collecting a random sample  of " & n & " " & _
        obsr &" you find a sample proportion of " & P & ". "

        Session("x")=x 
        Session("mea")=mea 
        Session("obsr")=obsr 
        Session("Pi")=Pi 
        Session("sig")=sig 
        Session("se")=se 
        Session("moe")=moe
        Session("n")=n 
        Session("z")=z
        Session("P")=P
        
        Session("CL") = CL
        
        Session("LowerB")=LowerB 
        Session("UpperB")=UpperB 

        Session("paragraph")=paragraph
        Session("NumAttempts") = numattempts
        Session("NumAttCon")=NumAttCon

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
  
 <h1 align="center"> Exercise on the Z Confidence Interval for the Population Proportion</h1>
   
   
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
      B. Round your answers to two decimal places.<br>
      C. You have <% response.write(session.timeout) %> minutes to complete your work before session
      timeout. After that your work <br> will not be graded correctly. </p>
   </td></tr></table>

 <table align="center" width="600">  
   

   <%
         PrHtml("<tr><td>" & "1. Z table value = <input  name=" & chr(34) & "AnsZ" & chr(34) & "></td></tr>" )    
         PrHtml("<tr><td>" & "2. Margin of Error = <input  name=" & chr(34) & "AnsMOE" & chr(34) & "></td></tr>" )
         
         PrHtml("<tr><td>" )
         PrHtml("3: You estimate with " & cl*100 & "% confidence that the population proportion falls between <br>")
         PrHtml("the lower value of " & " <input  name=" & chr(34) & "AnsLowerBound" & chr(34) & ">" )
         PrHtml("<br> and the upper value of " & " <input  name=" & chr(34) & "AnsUpperBound" & chr(34) & ">" )
         PrHtml("<br> (If you get Question 3 correct, you will be asked an additional question)</td></tr>" )
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


