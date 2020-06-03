<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>
<form action=ZCIforPiCon.asp method=POST enctype="application/x-www-form-urlencoded">
<h1 align="center">Confidence Interval for Population proportion Exercise </h1>
<br>
 <%   
   AnsZ=request.form("AnsZ")
   AnsMOE=request.form("AnsMOE")
   AnsLower=request.form("AnsLowerBound")
   AnsUpper=request.form("AnsUpperBound")
   
   session("AnsZ")=AnsZ
   session("AnsMOE")=AnsMoe


   x =   Session("x")	  
   mea = Session("mea")	
   obsr = Session("obsr")	
   Pi = Session("Pi")	
   sig = Session("sig")	
   n = Session("n")
   SE = Session("SE")	
   Z = Session("Z")
   MOE = session("moe")
   'prhtml("moe=" & moe)
   P=Session("P")

   CL = Session("CL")	
   
   UpperB=Session("UpperB") 
   LowerB=Session("LowerB")
   paragraph= Session("paragraph")


    name=request.form("Name")
    session("Name")=replace(name,",","-")
    session("Name")=replace(name," ","-")     

   NumAttempts = session("NumAttempts")
   NumAttempts = numattempts+1
   session("Numattempts") = numattempts

   prhtml("<p align=" & chr(34) & "center" & chr(34) & "> Number of attempts = " & numattempts & "<p>")
   
    sub PrHtml(strHtml)
       response.write(strhtml)
    end sub
   
   %>
   <p align="center"> <b><%response.write("Your name is: " & name)%></b> </p>
   <table align="center" border="1" width="600">
     <tr>
       <td align="left" style="word-wrap: break-word"> 
        <% if len(trim(paragraph))=0 then
            prhtml("<b> Error: The information within the page has been lost. Check to see if <br>" & _
                   "you have exceeded the time limit or perhaps your browswer security <br>" & _
                   "does not allow cookies. <br> <br> Program will not grade answers correctly.</b><br>")
           else
            prhtml(paragraph & "<br>") 
           end if
         %>
       </td>
     </tr>
     </table>	
  
   <br>

      
   <%
    Prhtml("<table align=" & chr(34) & "center" & chr(34) & "><tr><td>")

    prhtml("<br>In case of an error, if you are using Internet Explorer, then " & _
              "go Back to the previous page <br> and try again. The program will " & _
              "keep track of the number of attempts.<br><br>")
 
       prhtml("</td></tr><tr><td>")  

    Prhtml("<p align=" & chr(34) & "center" & chr(34) & ">Identify: This is asking you to estimate a population proportion.</p>")
    Prhtml("<p align=" & chr(34) & "center" & chr(34) & ">Calculate the sample standard error: First, find the value of P(1-P)/n = " & _
           p & "(1-" & p & ")/" & n & " = " & round(p*(1.-p)/n,8) & "<br> where P is the sample proportion and n is the sample size." & _
           "<br>The sample standard error is the square root of " & round(p*(1.-p)/n,8) & " = " & se & "</p>")
      

  'Question 1
    PrHtml("Question 1: ") 
      if isnumeric(ansz) = true then
    
        If abs(ansz-z)<=0.01  then
         
           PrHtml("Correct: The z table value is " & z & "<br>")
           
         Else
            Prhtml("Incorrect: Your answer of " & ansz & " is wrong. <br>")
         end if
      elseif len(ansz) =0 then
         prhtml("You did not enter an answer to Question 1. <br>")
      else
            Prhtml("Your answer of " & ansz & " is not a number.<br>")
      end if
     prhtml("<br></td></tr><tr><td>")

    PrHtml("Question 2: ") 
  
      if isnumeric(ansMOE) = true then
    
        If abs(ansMoe-moe)<=.015 then
         
           PrHtml("Correct: " & cl*100 & "% of the time, the largest error you would expect when using your sample"  & _
                   " <br> proportion to estimate your population proportion is " & MOE & ".<br>")
           
         Else
            Prhtml("Incorrect: Your answer of " & ansmoe & " is wrong. <br>")
         end if
      elseif len(ansmoe) =0 then
         prhtml("You did not enter an answer to Question 2. <br> ")
      else
            Prhtml("Your answer of " & ansmoe & " is not a number.<br>")
      end if
     prhtml("<br></td></tr><tr><td>")
    
     PrHtml("Question 3a: Lower Value: ") 
      Question3a ="Incorrect"
      if isnumeric(ansLower) = true then
    
        If abs(ansLower-lowerb) < 0.01*lowerb then
          
           Question3a ="Correct"
           PrHtml("Correct: " & round(lowerb,2) & " is the correct lower bound. <br>")
  
         Else
            Prhtml("Incorrect: Your answer for the lower value:  " & ansLower & " is wrong. <br>")
         end if
      elseif len(anslower) =0 then
         prhtml("You did not enter a lower value answer to Question 3 . <br>")
      else
            Prhtml("Your answer of " & anslower & " for the lower value is not a number.<br>")
      end if
     
     PrHtml("Question 3b: Upper Value: ") 
      Question3b = "Incorrect"

      if isnumeric(ansUpper) = true then
    
        If abs(ansupper-upperb) < 0.01*upperb then
          
           Question3b = "Correct"
           PrHtml("Correct: " & round(upperb,2) & " is the correct upper bound. <br>")
  
         Else
            Prhtml("Incorrect: Your answer for the upper value:  " & ansupper & " is wrong. <br>")
         end if
      elseif len(ansupper) =0 then
         prhtml("You did not enter an upper value answer to Question 3 . <br>")
      else
            Prhtml("Your answer of " & ansupper & " for the upper value is not a number.<br>")
      end if

     prhtml("<br></td></tr><tr><td>")

     if (Question3a = "Correct") and (Question3b = "Correct") then
       Prhtml("<b>Followup question: </b> Which of the following is the best interpretation of your results?<br><br>")


      call MultipleChoice(x, obsr, LowerB, UpperB, P, mea)

     end if
      
     prhtml("</td></tr></table>")

   
   Sub MultipleChoice(x, obsr, LowerB, UpperB, P, mea)
      Dim Original()
      Dim Corrections()
      Dim ranked()
      Dim OldList()
      Dim NewList()
      Dim OldPosition()
      Dim NewCorrections()
      Dim NumRows
      NumRows = 5
      ReDim Original(NumRows)
      ReDim ranked(NumRows)
      ReDim OldList(NumRows)
      ReDim NewList(NumRows)
      ReDim OldPosition(NumRows)
      ReDim Corrections(NumRows)
      ReDim NewCorrections(NumRows)
     randomize
     for i =1 to numrows
        Original(i) = rnd()
     next
  
     OldList(1) ="With " & cl*100 & "% confidence, we can say that the proportion of all " & _
          obsr & " that are " & x & " falls between " & LowerB & " and " & UpperB & "."
     OldList(2)="With " & cl*100 & "% confidence, we can say that all " & obsr & " " & x & _
           " falls between " & LowerB & " and " & UpperB & "."
     OldList(3)="With " & cl*100 & "% confidence, we can say that the proportion of the sample of " & _
          obsr & " that are " & x & " falls between " & LowerB & " and " & UpperB & "."
     OldList(4)="With " & cl*100 & "% confidence, we can say that the proportion  " & _
        " of all " & obsr & " that are " & x & " equals " & P & "."
     OldList(5)="The proportion of all " & obsr & " that are " & x & " equals " & P & "."
   
     Corrections(1) =" This is correct. " 
     Corrections(2)="Incorrect, you were asked to estimate the value of the population proportion, " & _
            "not the value of the " & x & "."
     Corrections(3)="Incorrect, you were asked to estimate the population proportion, not the  sample proportion." & _
            " Plus you are 100% certain the sample proportion is in the interval."
     Corrections(4)="Incorrect, if you use this form, you must include the margin of error of your estimate."
     Corrections(5)="Incorrect, the value of " & P & " is your sample proportion and in almost all " & _
             "cases does not equal the population proportion."

   
     Call RankArray(Original, ranked)

   For i = 1 To NumRows 

     OldPosition(i ) = ranked(i )
     
     NewList(i ) = chr(64+i) & ". " & OldList(ranked(i ))
     NewCorrections(i ) = Corrections(ranked(i ))
     
     'prhtml(NewList(i ))
     'prhtml("C4" & NewCorrections(i ) )
    next
     
    prhtml("<input type=" & Chr(34) & "radio"        & Chr(34) & _
                " checked=" & Chr(34) & "checked"      & Chr(34) & _
                  " name=" & Chr(34) & "Conclusion"   & Chr(34) & _
                  " value=" & Chr(34) & 99 & Chr(34) & "><br>")
               
    For i = 1 To 5
        prhtml("<input type=" & Chr(34) & "radio"      & Chr(34) & _
                     " name=" & Chr(34) & "Conclusion" & Chr(34) & _
                     " value=" & Chr(34) & i    & Chr(34) & ">" & newlist(i) & "<br>")
     Next
   
    Session("NewList")=Newlist
    Session("NewCorrections") = NewCorrections



   Prhtml("<p    align=" & Chr(34) & "center"                             & Chr(34) & ">" & _
          "<input type=" & Chr(34) & "submit"                             & Chr(34) & _
               " value=" & Chr(34) & "Please click here to check conclusion" & Chr(34) & _
                " name=" & Chr(34) & "END"                                & Chr(34) & "</p>")


   End sub

Sub RankArray(Old, ranks)
    Dim copy()
    NumRows = UBound(Old)
    ReDim copy(NumRows)
    
    For i = 1 To NumRows
       copy(i) = Old(i)
       
    Next 
    'sort the values and place sorted values in copy
    For i = 1 To NumRows
      Min = copy(i)
      Rank = ranks(i)
      
      For j = i + 1 To NumRows
         If Min > copy(j) Then
            Min = copy(j)
            copy(j) = copy(i)
            copy(i) = Min
         End If
      Next 
    Next 
    'go through each of the values and determines its position in
    'ranked array
    For i = 1 To NumRows
      For j = 1 To NumRows
         If Old(i) = copy(j) Then
             ranks(i) = j
         End If
      Next 
    Next 
   
End Sub

    function ChkAns(a,b)
      if abs(a-b) < 0.0015 then
        chkans=1
      else
        chkans=0
      end if
    end function
    %>

<p  align="center">
Click <a  href="zConfIntForPiQues.asp">here</a> for a new question
</p>
</body>

</html>