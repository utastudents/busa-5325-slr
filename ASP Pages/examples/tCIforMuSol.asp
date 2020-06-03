<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>

<form action=tCIforMuCon.asp method=POST enctype="application/x-www-form-urlencoded">
<h1 align="center">t Confidence Interval for Population Mean Exercise </h1>
<br>
 <%   
   Anst=request.form("Anst")
   AnsMOE=request.form("AnsMOE")
   AnsLower=request.form("AnsLowerBound")
   AnsUpper=request.form("AnsUpperBound")
  
   session("Anst")=Anst
   session("AnsMOE")=AnsMoe

   x =   Session("x")	  
   mea = Session("mea")	
   obsr = Session("obsr")	
   mu = Session("mu")	
   s = Session("s")	
   n = Session("n")
   SE = Session("SE")	
   t = Session("t")
   MOE = round(t*se,2)
    xbar=Session("xbar")
   
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

    Prhtml("<p align=" & chr(34) & "center" & chr(34) & ">Identify: This is asking you to estimate a population mean and the population standard deviation is unknown.</p>")
    Prhtml("<p align=" & chr(34) & "center" & chr(34) & ">Calculate the sample standard error: This will be " & _
            s & " divided by the square root of " & n & " = " & se & "</p>")
      

  'Question 1
    PrHtml("Question 1: ") 
      session("Question1")="Incorrect"
      if isnumeric(anst) = true then
    
        If abs(anst-round(t,2))<=0.01  then
         
           PrHtml("Correct: The t table value is " & t & "<br>")
           session("Question1") = "Correct"
           
         Else
            Prhtml("Incorrect: Your answer of " & anst & " is wrong. <br>")
         end if
      elseif len(anst) =0 then
         prhtml("You did not enter an answer to Question 1. <br>")
      else
            Prhtml("Your answer of " & anst & " is not a number.<br>")
      end if
     prhtml("<br></td></tr><tr><td>")

    PrHtml("Question 2: ") 
      session("Question2")="Incorrect"
      if isnumeric(ansMOE) = true then
    
        If abs(ansMoe-moe)<.01*moe then
         
           PrHtml("Correct: " & cl*100 & "% of the time, the largest error you would expect when using your sample"  & _
                   " average to estimate your population average is " & MOE & " " & mea & ".<br>")
           session("Question2") = "Correct"
           
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
      session("Question3a")="Incorrect"
      if isnumeric(ansLower) = true then
    
        If abs(ansLower-lowerb) < 0.01*lowerb then
          
         
           PrHtml("Correct: " & round(lowerb,2) & " is the correct lower bound. <br>")
           Question3a = "Correct"
           session("Questiona")="Correct"
         Else
            Prhtml("Incorrect: Your answer for the lower value:  " & ansLower & " is wrong. <br>")
         end if
      elseif len(anslower) =0 then
         prhtml("You did not enter a lower value answer to Question 3 . <br>")
      else
            Prhtml("Your answer of " & anslower & " for the lower value is not a number.<br>")
      end if
     
     PrHtml("Question 3b: Upper Value: ") 
      session("Question3b")="Incorrect"
      if isnumeric(ansUpper) = true then
    
        If abs(ansupper-upperb) < 0.01*upperb then
          
         
           PrHtml("Correct: " & round(upperb,2) & " is the correct upper bound. <br>")
           Question3b = "Correct"
           session("Question3b")="Correct"
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
       Prhtml("<b>Followup question: </b> Which of the following is the best interpretation of your      results?<br><br>")


     call MultipleChoice(x, obsr, LowerB, UpperB, xbar, mea)

     end if
     prhtml("</td></tr></table>")

Sub MultipleChoice(x, obsr, LowerB, UpperB, xbar, mea)
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
   
     OldList(1) ="With " & cl*100 & "% confidence, we can say that the mean " & x & _
         " of all " & obsr & " falls between " & LowerB & " and " & UpperB & " " & mea & "."
     OldList(2)="With " & cl*100 & "% confidence, we can say that the " & x & " of all " & _
           obsr & " falls between " & LowerB & " and " & UpperB & " " & mea & "."
     OldList(3)="With " & cl*100 & "% confidence, we can say that the average " & x & _
         " of the sample of " & obsr & " falls between " & LowerB & " and " & UpperB & " " & mea & "."
     OldList(4)="With " & cl*100 & "% confidence, we can say that the mean  " & x & _
        " of all " & obsr & " equals " & xbar & " " & mea & "."
     OldList(5)="The mean  " & x & " of all " & obsr & " equals " & xbar & " " & mea & "."
   
     Corrections(1) =" This is correct. " 
     Corrections(2)="Incorrect, you were asked to estimate the value of the population mean, " & _
            "not the value of the " & x & "."
     Corrections(3)="Incorrect, you were asked to estimate the population mean, not the  sample mean." & _
            " Plus you are 100% certain the sample mean is in the interval."
     Corrections(4)="Incorrect, if you use this form, you must include the margin of error of your estimate."
     Corrections(5)="Incorrect, the value of " & xbar & " is your sample mean and in almost all " & _
             "cases does not equal the population mean."

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
Click <a href="tConfIntForMuQues.asp">here</a> for a new question
</body>

</html>