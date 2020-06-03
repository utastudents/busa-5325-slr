<html>
<head>

</head>


<body>
<form action=EmpiricalRuleSol.asp method=POST enctype="application/x-www-form-urlencoded">
 <%   dim Question()
    redim Question(4)
NumAttempts = 0
    session.timeout=30
   'Creation of scenarios
       variable = Array("age", "mpg ", "cost", "sales", "price", "height")
        measure = Array("years", "miles", "dollars", "dollars", "dollars", "inches")
        observation = Array("buyers", "cars", "trips", "LCD panels", "tickets", "students")
        mean = Array(41, 27, 115, 632, 23, 35)
        st_dev = Array(7, 5, 15, 66, 6, 1)
        
        Randomize
             Scenario = Round(rnd() * 6 + 0.5, 0) - 1
             
    'This scenarios values
        x = variable(Scenario)
        mea = measure(Scenario)
        obsr = observation(Scenario)
        mu = mean(Scenario)
        sig = st_dev(Scenario)
       
    'Populating the empirical rule for this scenario
        choice = Array(1, 2, 3, 4, 5, 6, 7)
        bounds = Array(mu - 3 * sig, mu - 2 * sig, mu - sig, mu, mu + sig, mu + 2 * sig, mu + 3 * sig)
        Prob = Array(0, 0.025, 0.135, 0.34, 0.34, 0.135, 0.025)
        prless = Array(0, 0.025, 0.16, 0.50, 0.84, 0.975, 1.00)
        prGreater = Array(1, 0.975, 0.84, 0.5, 0.16, 0.025, 0)
        
    'Determing a random < range and its prob
        choice = Round(rnd() * 7 + 0.5, 0) - 1
        LessThanValue = bounds(choice)
        LessThanProb = prless(choice)
        Question(1) = "1. What is the probability of finding " & obsr & " with a value of " & x & _
                  " less than " & LessThanValue & " " & mea & "?"
        
    'Determing a random > range and its prob
        choice = Round(rnd() * 7 + 0.5, 0) - 1
        GrThanValue = bounds(choice)
        GrThanProb = prGreater(choice)
        Question(2) = "2. What is the probability of finding " & obsr & " with a value of " & x & _
                   " greater than " & GrThanValue & " " & mea & "?"
        
    'Determining a random range and then finding its prob
        choiceupper = RandBetween(1, 6)
        choicelower = RandBetween(0, choiceupper - 1)
        uppervalue = bounds(choiceupper)
        lowervalue = bounds(choicelower)
        BetweenProb = prless(choiceupper) - prless(choicelower)
        Question(3) = "3. What is the probability of finding " & obsr & " with a value of " & x & _
            " between " & lowervalue & " and " & uppervalue & " " & mea & "?"

    'Determining a determining a percentile and then finding its value of x
        choice = Round(rnd() * 7 + 0.5, 0) - 1
        Percentile = prless(choice) * 100
        X_for_pctile = bounds(choice)
        Question(4) = "4. What value of " & x & " falls on the " & Percentile & " percentile?"
  
     paragraph = "Suppose the " & x & " of " & obsr & " were collected and the data follows " & _
     "the bell curve (a normal distribution). The " & obsr & " had an average " & x & _
     " of " & mu & " " & mea & " and a standard deviation of " & sig & _
     ". The ranges and percentages become"


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
        Session("X_for_pctile")=X_for_pctile 
        Session("paragraph")=paragraph
        Session("NumAttempts") = numattempts

        Name=""
        if Session("Name")<>"" then name=replace(session("Name")," ","-")

    
    Function RandBetween(a, b)
      RandBetween = Round(Rnd() * (b - a + 1) + 0.5, 0) + a - 1
    End Function
    
    sub PrHtml(strHTML)
       response.write(strHTML)
    end sub

    
 %>
   <h1 align="center"> Empirical Rule Exercise</h1>
   <p align="center">In these empirical rule examples, it is assumed that 100% fall within three sigmas of mu. <br> (99.7% is rounded up to 100%)</p>
   
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

   <table align="center" border="1" width="300">
     <tr>
       <td> <b>Range</b></td>
       <td> <b>Probability</b></td>
       <td> <b>Cumulative Prob</b></td>
     </tr>
    <%
    
    for i=1 to 6
       prHTML("<tr>")
        prhtml("<td>" & bounds(i-1) & " to " & bounds(i) & "</td>")
        prhtml("<td align=" & chr(34) & "center" & chr(34) & ">" & prob(i) & "</td>")
        Prhtml("<td align=" & chr(34) & "center" & chr(34) & ">" & prless(i) & "</td>")
       prhtml("</tr>")
    next

   %>
   </table><br>
 
   <table align="center" border="1"  width="600"> 
    <tr><td>
      <b> Directions: Fill in the blanks and click the button to have your answers graded.</b><br>
      A. For questions 1-3 enter a number between 0 and 1. <br>
      B. You have <% response.write(session.timeout) %> minutes to complete your work before session
      timeout. After that your work <br> will not be graded correctly. </p>
    </td></tr></table> <br>
 
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

</body>

</html>