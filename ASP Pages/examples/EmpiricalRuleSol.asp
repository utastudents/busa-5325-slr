<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>
<h1 align="center"> Empirical Rule Exercise </h1>
<br>

  
 <%   
   Ans1=request.form("AnsQues1")
   Ans2=request.form("AnsQues2")
   Ans3=request.form("AnsQues3")
   Ans4=request.form("AnsQues4")

   x =   Session("x")	  
   mea = Session("mea")	
   obsr = Session("obsr")	
   mu = Session("mu")	
   sig = Session("sig")	
   bounds = Session("bounds")
   Prob = Session("Prob")	
   prless = Session("prless")
   prGreater = Session("prGreater")
   LessThanProb = Session("LessThanProb")
   GrThanProb = Session("GrThanProb")
   BetweenProb= Session("BetweenProb")
   LessThanValue=Session("LessThanValue") 
   GrThanValue =Session("GrThanValue")
   UpperValue=Session("UpperValue") 
   LowerValue=Session("LowerValue")
   Percentile=Session("percentile")
   X_for_pctile = Session("X_for_pctile")
   paragraph= Session("paragraph")


   NumAttempts = session("NumAttempts")
   NumAttempts = numattempts+1
   session("Numattempts") = numattempts


    name=request.form("Name")
    session("Name")=replace(name,",","-")
    session("Name")=replace(name," ","-")     

   prhtml("<p align=" & chr(34) & "center" & chr(34) & "> Number of attempts = " & numattempts & "<p>")
   
    sub PrHtml(strHtml)
       response.write(strhtml)
    end sub
   
   %>
   <p align="center"> <b><%response.write("Your name is: " & name)%></b> </p>

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
       <td> <b>Percent</b></td>
       <td> <b>Cumulative Pct</b></td>
     </tr>
    
    <%
    
    for i=1 to 6
       prHTML("<tr>")
        prhtml("<td>" & bounds(i-1) & " to " & bounds(i) & "</td>")
        prhtml("<td>" & prob(i)*100 & "%</td>")
        Prhtml("<td>" & prless(i)*100 & "%</td>")
       prhtml("</tr>")
    next

    %>

   </table><br>
   
   <%
    Prhtml("<table align=" & chr(34) & "center" & chr(34) & "><tr><td>")

    prhtml("<br>In case of an error, if you are using Internet Explorer, then " & _
              "go Back to the previous page <br> and try again. The program will " & _
              "keep track of the number of attempts.<br><br>")
 
       prhtml("</td></tr><tr><td>")  
   'Question 1
    PrHtml("Question 1: ") 
      if isnumeric(ans1) = true then
    
        If chkans(ans1,LessThanProb) =1 then
         
           PrHtml("Correct: The probability of finding " & obsr & " with a value of " & x & _
                  " less than " & LessThanValue & " " & mea & " is " & LessThanProb & "<br>")
           
         Else
            Prhtml("Incorrect: Your answer of " & ans1 & " is wrong. <br>")
         end if
      elseif len(ans1) =0 then
         prhtml("You did not enter an answer to Question 1. <br>")
      else
            Prhtml("Your answer of " & ans1 & " is not a number between 0 and 1.<br>")
      end if
     prhtml("<br></td></tr><tr><td>")

    PrHtml("Question 2: ") 
      if isnumeric(ans2) = true then
    
        If chkans(ans2,GrThanProb) =1 then
         
           PrHtml("Correct: The probability of finding " & obsr & " with a value of " & x & _
                   " greater than " & GrThanValue & " " & mea & " is " & GrThanProb & "<br>")
           
         Else
            Prhtml("Incorrect: Your answer of " & ans2 & " is wrong. <br>")
         end if
      elseif len(ans2) =0 then
         prhtml("You did not enter an answer to Question 2. <br> ")
      else
            Prhtml("Your answer of " & ans2 & " is not a number between 0 and 1.<br>")
      end if
     prhtml("<br></td></tr><tr><td>")
    
     PrHtml("Question 3: ") 
      if isnumeric(ans3) = true then
    
        If chkans(ans3,BetweenProb) =1 then
         
           PrHtml("Correct: The probability of finding " & obsr & " with a value of " & x & _
            " between " & lowervalue & " and " & uppervalue & " " & mea & " is " & BetweenProb & "<br>")
           
         Else
            Prhtml("Incorrect: Your answer of " & ans3 & " is wrong. <br>")
         end if
      elseif len(ans3) =0 then
         prhtml("You did not enter an answer to Question 3. <br> ")
      else
            Prhtml("Your answer of " & ans3 & " is not a number between 0 and 1. <br>")
      end if
     prhtml("<br></td></tr><tr><td>")
     
     PrHtml("Question 4: ") 
      if isnumeric(ans4) = true then
    
        If chkans(ans4,X_for_pctile) =1 then
         
           PrHtml("Correct: The value of " & X_for_pctile & " falls on the " & Percentile & " percentile." & "<br>")
  
         Else
            Prhtml("Incorrect: Your answer of " & ans4 & " is wrong. <br>")
         end if
      elseif len(ans4) =0 then
         prhtml("You did not enter an answer to Question 4. <br>")
      else
            Prhtml("Your answer of " & ans4 & " is not a number.<br>")
      end if
     
     prhtml("<br></td></tr><tr><td>")
     prhtml("</td></tr></table>")

   function ChkAns(a,b)
      if abs(a-b) < 0.0015 then
        chkans=1
      else
        chkans=0
      end if
    end function
    %>
Click <a href="EmpiricalRuleques.asp">here</a> for a new question
</body>

</html>