<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>
 <%   dim sumX, i, sngmean
 <%   
   Places = Array( _
             "A disease is striking a small country where  ", _
             "In a small town ", _
             "Among the students at a college ", _
             "Among car buyers responding to a survey ", _
             "Of all the transactions at a company ", _
             "In a population of customers" _
             )
    Thing = array( _
            "person", _
            "person", _
            "student", _
            "car buyer", _
            "transaction", _
            "customer" _ 
            )
    Things = Array( _
            "people", _
            "people", _
            "students ", _
            "car buyers", _
            "transactions", _
            "customers" _
            )
    Cat1Lvl1 = Array( _
            "have the disease", _
            "watch football", _
            "use facebook", _
            "buy extended warranties", _
            "are over $80", _
            "buy product A" _
            )
    Cat1Lvl2=array( _
            "do not have the disease", _
            "do not watch football", _
            "do not use facebook", _
            "do not buy extended warranties", _
            "are $80 or less", _
            "do not buy product A" _
            )
    Cat1Lvl1Single=array( _
            "does have the disease", _
            "does watch football", _
            "does use facebook", _
            "does buy an extended warranty", _
            "is over $80", _
            "does buy product A" _
            )
    Cat2Lvl1=array( _
           "have a blood test that comes back positive", _
            "drink beer", _
            "use twitter", _
            "will have a major repair after 5 years", _
            "are taxable ", _
            "buy product B" _
            )

    Cat2Lvl2 = Array( _
          "have a blood test that comes back negative", _
            "drink wine", _
            "do not use twitter", _
            "do not have a major repair", _
            "are not taxable", _
            "do not buy product B" _
             )
     Cat2Lvl1Single=array( _
           "has a blood test that comes back positive", _
            "drinks beer", _
            "uses twitter", _
            "will have a major repair within 5 years", _
            "is taxable", _
            "buys product B" _
             )
    NumAttempts=0
   randomize
     chosen = round(rnd()*6+.5,0)-1
     place = places(chosen)
     obsr = thing(chosen)
     obsrs = things(chosen)
     Cat1L1 = Cat1Lvl1(chosen)
     Cat1L2 = Cat1Lvl2(chosen)
     Cat1L1obsr = Cat1Lvl1Single(chosen)
     Cat2L1 = Cat2Lvl1(chosen)
     Cat2L2 = Cat2Lvl2(chosen)
     Cat2L1obsr = Cat2Lvl1Single(chosen)

     Margin1 = 90 - round(rnd()*80+.5)
     Cond1 = 90 - round(rnd()*80+.5)
     Cond2 = 90 - round(rnd()*80+.5)

     Sentence1 = place & " " & margin1 & "% " & cat1l1 & ". "
     Sentence2 = "Of the " & obsrs & " that " & cat1l1 & ", " & cond1 & "% " & cat2l1 & ". "
     Sentence3 = "While among those " & obsrs & " that " & cat1l2 & ", " & cond2 & "% " & cat2l1 & ". "
     
     Question = "Given that a randomly selected " & obsr & " " & cat2l1obsr & ", " & _
                 "what is the chance that the " & obsr & " " & cat1l1obsr & "?"
     Paragraph = sentence1 & " " & sentence2 & " " & sentence3 & " " & Question



   sumX=0.0 
  


   Numerator = margin1*cond1/100
   Denominator =(numerator+(100-Margin1)*Cond2/100)
   Answer = csng(numerator)/csng(denominator)


   %>
<h1 align="center"> Bayes Example</h1>


<br> <br>
  
  
<table align="center" border="1" width="600">

   <tr>
     <td align="left" style="word-wrap: break-word"> 
       <%response.write(paragraph)%>  <br>

     </td>
   </tr>
   <tr>
     <td>
      <b> Solution </b>
      	        
     </td>
    </tr>	
  <tr>
     <td>
      <%Response.write("a. Determine the percent in the first major category, those that " & Cat1L1  &"."    )%><br><br>
      <%Response.write("To start lets assume there are 10,000 " & obsrs & ".")%><br>
      <%Response.write( Margin1 & "% or " & Margin1*100 & " " & obsrs & " " & Cat1L1)%><br>
      <%Response.write("therefore")%><br>
      <%Response.write( (100-Margin1) & "% or " & (100-Margin1)*100 & " "  & Cat1L2 & ".")%><br><br>
     </td>
    </tr>
   <tr>
     <td>
      <%Response.write("b. Of the levels of the first major category, what percent of population belong to the second?" )%><br>
      <br>
      <%Response.write(sentence2)%> <br>
      and <br>
      <%Response.write(sentence3)%> <br>
    	        
     </td>
    </tr>
    
    <tr>
     <td>
      <%Response.write("c. Determine the percent of all " & obsrs & " that have the combined characteristics." )%><br>
      <br>
      <%Response.write("(" & Cond1 & "% of " & Margin1*100 & ") or  " & Cond1*Margin1 & " " & obsrs & " both " & Cat1L1 & " and " & Cat2L1)%> <br>
      while <br>
      <%Response.write("(" & Cond2 & "% of " & (100-Margin1)*100 & ") or  " & (100-Margin1)*Cond2 & " " & obsrs & " both " & Cat1L2 & " and " & Cat2L1)%> <br>
    	        
     </td>
    </tr>

    <tr>
     <td>
      <%Response.write("d. Determine what is the total percent of all " & obsr & " that " & Cat2L1 )%><br>
      <br>
      <%Response.write("Of the total number that " & Cat2L1 & " some " & Cat1L1 & " and some " & Cat1L2)%> <br>
      <%Response.write("This is the sum of the previous two values: " & Cond1*Margin1 & " and " & Cond2*(100-Margin1) & ".")%> <br>
      <%Response.write("Therefore " & (Cond1*Margin1+(100-Margin1)*Cond2) & " " & obsrs & " "  & Cat2L1)%> <br>
      	        
     </td>
    </tr>

    <tr>
     <td>
      <%Response.write("e. To find your answer to the problem, determine among all " & obsr & " that " & Cat2L1 & " what proportion " & Cat1L1)%><br><br>
      
      <%Response.write("Of the " & (Margin1*Cond1+(100-Margin1)*Cond2) & "  that " & Cat2L1 & ", " & Margin1*Cond1 & " " & Cat1L1)%> <br>
      <%Response.write("Calculate the fraction " & Cond1*Margin1 & " divided by " & ((100-Margin1)*Cond2+Margin1*Cond1) & "")%> <br>
      <%Response.write("The answer to the exercise is " & Numerator*100 & "/" & Denominator*100 & " = " & ROUND(Numerator/Denominator,4)*100 & "% (rounding to two decimal places.)")%> <br>
      	        
     </td>
    </tr>

</table>
Click <a href="<%= request.servervariables("script_name") %>">here</a> for a new example
</body>

</html>