
<HTML>
<HEAD><TITLE>t CI on Mu</TITLE>


</HEAD>
<BODY>
<form action=tCIforMuSol.asp method=POST enctype="application/x-www-form-urlencoded">
<%
   Dim Question()

    ReDim Question(4)
    NumAttempts = 0
    NumAttCon = 0

    Session.timeout=30
   'Creation of scenarios
       variable = Array("age", "mpg ", "cost", "sales", "price", "height")
        measure = Array("years", "miles", "dollars", "dollars", "dollars", "inches")
        observation = Array("buyers", "cars", "trips", "LCD panels", "tickets", "students")
        CL_Levels = Array(0.80,0.90,0.95,0.98,0.99)
        Z_Values = Array(1.281551566, _
                         1.644853627, _
                         1.959963985, _
                         2.326347874, _
                         2.575829304)

   Randomize 
       mean = Array(Round(randbetween(30, 50), 0), _
                     Round(randbetween(15, 35), 0), _
                     Round(randbetween(100, 130), 0), _
                     Round(randbetween(230, 900), 0), _
                     Round(randbetween(30, 130), 0), _
                     Round(randbetween(40, 60), 0))
       
        st_dev = Array(Round(randbetween(1, 7), 1), _
                     Round(randbetween(1, 5), 1), _
                     Round(randbetween(10, 20), 1), _
                     Round(randbetween(30, 50), 1), _
                     Round(randbetween(2, 8), 1), _
                     Round(randbetween(1, 5), 1))
       
     
    
         Scenario = Round(Rnd() * 6 + 0.5, 0) - 1
         ChosenProbIndex = round(rnd()*5+.5,0)-1
         Z = z_Values(chosenProbIndex)
         CL = Cl_Levels(chosenProbIndex)

    'This scenarios values
        x = variable(Scenario)
        mea = measure(Scenario)
        obsr = observation(Scenario)
        xbar = mean(Scenario)
        s = st_dev(Scenario)
    'Determine a sample size between 8 and 100
        n = Round(randbetween(8, 100), 0)
       
     'Determine the sample standard error value
        SE = round(s/sqr(n),8)

    
    'Determining a determining middle proportion and then finding the two xbar values 
        

        'Find area less than the upper range
        UpperPercentile = 1.-(1-Cl)/2   
        df = n-1
        'Prhtml("df = " & df & ", z = " & z & ", upperpercentile = " & upperpercentile & "<br>")

        t =  round(PPF(UpperPercentile, df, z),3)

        LowerB = round(xbar - t * SE,2)
        UpperB = round(xbar + t * SE,2)
        'prhtml("<p> t = " & t & " </p>")
        'prhtml("<p> Lower Bound = " & LowerB & ", " & "Upper Bound = " & UpperB & "</p>")

       

 
     paragraph = "Estimate the mean "& x &" of all " & obsr & " with " & CL*100 & "% " & _
       "confidence. After collecting a random sample  of " & n & " " & _
        obsr &" you find a sample mean of " & xbar & " " & mea & " and a sample" & _
        " standard deviation of " & s & " " & mea & ". Assume the " & _
       "distribution of " & x & " for all " & obsr & " is normal."
 
        Session("x")=x 
        Session("mea")=mea 
        Session("obsr")=obsr 
        Session("mu")=mu 
        Session("s")=s 
        Session("se")=se 
        Session("n")=n 
        Session("t")=t
        Session("xbar")=xbar
        
        Session("CL") = CL
        
        Session("LowerB")=LowerB 
        Session("UpperB")=UpperB 

        Session("paragraph")=paragraph
        Session("NumAttempts") = numattempts
        Session("NumAttCon")=NumAttCon

        Name=""
        if Session("Name")<>"" then name=replace(session("Name")," ","-")
  
  Function PPF(P, NU, z)

      Pi = 3.14159265358979
      SQRT2 = 1.414213562
      B21 = 0.25
      B31 = 0.01041666666667
      B32 = 5.0
      B33 = 16.0
      B34 = 3.0
      B41 = 0.00260416666667
      B42 = 3.0
      B43 = 19.0
      B44 = 17.0
      B45 = -15.0
      B51 = 0.00001085069444
      B52 = 79.0
      B53 = 776.0
      B54 = 1482.0
      B55 = -1920.0
      B56 = -945.0
      IPR = 6
      DNU = NU
      DP = P
      MAXIT = 5
'
      If NU = 1 Then
'
'        TREAT THE NU = 1 (CAUCHY) CASE
'
         DARG = Pi * DP
         PPF = -Cos(DARG) / Sin(DARG)
      End If
       
      DPPFN = z
      D1 = DPPFN
      D3 = DPPFN ^ 3
      D5 = DPPFN ^ 5
      D7 = DPPFN ^ 7
      D9 = DPPFN ^ 9
      TERM1 = D1
      TERM2 = B21 * (D3 + D1) / DNU
      TERM3 = B31 * (B32 * D5 + B33 * D3 + B34 * D1) / (DNU ^ 2)
      TERM4 = B41 * (B42 * D7 + B43 * D5 + B44 * D3 + B45 * D1) / (DNU ^ 3)
      TERM5 = B51 * (B52 * D9 + B53 * D7 + B54 * D5 + B55 * D3 + B56 * D1) / (DNU ^ 4)
      DPPF = TERM1 + TERM2 + TERM3 + TERM4 + TERM5
      PPF = DPPF
  
   End Function





 Function RandBetween(xl, xu)
      RandBetween = Round(Rnd() * (xu - xl), 2) + xl
    End Function
     

  sub PrHtml(strHtml)
     response.write(strhtml)
  end sub
%>
  
 <h1 align="center"> Exercise on the t Confidence Interval for the Population Mean</h1>
   
   
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
         PrHtml("<tr><td>" & "1. t table value = <input  name=" & chr(34) & "Anst" & chr(34) & "></td></tr>" )    
         PrHtml("<tr><td>" & "2. Margin of Error = <input  name=" & chr(34) & "AnsMOE" & chr(34) & "></td></tr>" )
         
         PrHtml("<tr><td>" )
         PrHtml("3: You estimate with " & cl*100 & "% confidence that the population mean falls between <br>")
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

