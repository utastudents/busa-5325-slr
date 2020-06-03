<html>
<head>

</head>


<body>
<form action=BayesSol.asp method=POST enctype="application/x-www-form-urlencoded">
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
    session.timeout=30

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
  


     Session("chosen")=Chosen	
     Session("place")= place 
     Session("obsr")= obsr 
     Session("obsrs")= obsrs 
     Session("Cat1L1")= Cat1L1 
     Session("Cat1L2")= Cat1L2 
     Session("Cat1L1obsr")= Cat1L1obsr 
     Session("Cat2L1")= Cat2L1 
     Session("Cat2L2")= Cat2L2 
     Session("Cat2L1obsr")= Cat2L1obsr 
     Session("Sentence2")=sentence2
     Session("Sentence3")=sentence3
     Session("Paragraph") = Paragraph
     Session("Margin1") =Margin1
     Session("Cond1")=Cond1
     Session("Cond2")=Cond2
     Session("NumAttempts") = numattempts

     Name=""
     if Session("Name")<>"" then name=replace(session("Name")," ","-")
 %><!--  -->
<h1 align="center"> Bayes Exercise</h1>


<table align="center" border="1" width="600">

   <tr>
     <td align="left" style="word-wrap: break-word"> 
       <%response.write(paragraph)%>  <br>

     </td>
   </tr>	
  <tr>
     <td>
      1. Enter your name, last name first
		<input style="width:350px" name="name" value=<%=Name%>>
      	        
     </td>
    </tr>
   <tr>
     <td><br>
      2. Enter your answer as a ratio, enter numerator and then denominator:              
		<input style="width:50px" name="AnsNum"  ><%response.write(" / ")%>
                <input style="width:50px" name = "AnsDen"><br>
      	        
     </td>
    </tr>
    <tr>
      <td> <br>
       Note: You have <% response.write(session.timeout) %> minutes to complete your work before session
       timeout. After that your work <br> will not be graded correctly. 
      </td>
    </tr> 
   
</table>
<br><br>
<p align="center"> <input type="submit" value="Please click here to check answers" name="END"></p>

<p align="center"> Click <a href="<%= request.servervariables("script_name") %>">here</a> for a new exercise</p>
</body>

</html>