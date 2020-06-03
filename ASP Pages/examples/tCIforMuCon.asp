<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>

<h1 align="center">t Confidence Interval for Population Mean Exercise </h1>
<br>
 <%   
   Ans4=request.form("Conclusion")
  
   Anst=session("Anst")
   AnsMoe=session("AnsMOE")
  

   x =   Session("x")	  
   mea = Session("mea")	
   obsr = Session("obsr")	
   mu = Session("mu")	
   sig = Session("sig")	
   n = Session("n")
   SE = Session("SE")	
   t = Session("t")
   MOE = round(t*se,2)
   xbar=Session("xbar")
   Newlist=Session("NewList")
   NewCorrections=Session("NewCorrections")
   
   CL = Session("CL")	
   
   UpperB=Session("UpperB") 
   LowerB=Session("LowerB")
   paragraph= Session("paragraph")


    name=session("Name")
   

   NumAttempts = session("NumAttempts")
   NumAttCon = session("NumAttCon")
   NumAttCon = numattCon+1
   session("NumattCon") = numattCon

   prhtml("<p align=" & chr(34) & "center" & chr(34) & "> Number of attempts at calculation= " & numattempts & "<p>")
   prhtml("<p align=" & chr(34) & "center" & chr(34) & "> Number of attempts at conclusion= " & numattcon & "<p>")
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

      
   <%
    Prhtml("<table align=" & chr(34) & "center" & chr(34) & "><tr><td>")

    prhtml("<br>In case of an error, if you are using Internet Explorer, then " & _
              "go Back to the previous page <br> and try again. The program will " & _
              "keep track of the number of attempts.<br><br>")
 
       prhtml("</td></tr><tr><td>")  

    Prhtml("<p align=" & chr(34) & "center" & chr(34) & ">Identify: This is asking you to estimate a population mean.</p>")
    Prhtml("<p align=" & chr(34) & "center" & chr(34) & ">Calculate the Standard error: This will be " & _
            sig & " divided by the square root of " & n & " = " & se & "</p>")
      

  'Question 1
    PrHtml("Question 1: ") 
      if isnumeric(Anst) = true then
    
        If abs(anst-round(t,2))<=0.01  then
         
           PrHtml("Correct: The t table value is " & t & "<br>")
           
         Else
            Prhtml("Incorrect: Your answer of " & Anst & " is wrong. <br>")
         end if
      elseif len(Anst) =0 then
         prhtml("You did not enter an answer to Question 1. <br>")
      else
            Prhtml("Your answer of " & Anst & " is not a number.<br>")
      end if
     prhtml("<br></td></tr><tr><td>")

    PrHtml("Question 2: ") 
      if isnumeric(ansMOE) = true then
    
        If abs(ansMoe-moe)<.01*moe then
         
           PrHtml("Correct: " & cl*100 & "% of the time, the largest error you would expect when using your sample"  & _
                   " average to estimate your population average is " & MOE & " " & mea & ".<br>")
           
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
      
          
           Question3a ="Correct"
           PrHtml("Correct: " & round(lowerb,2) & " is the correct lower bound. <br>")
  
    
     
     PrHtml("Question 3b: Upper Value: ") 

          
           Question3b = "Correct"
           PrHtml("Correct: " & round(upperb,2) & " is the correct upper bound. <br>")

     prhtml("<br></td></tr><tr><td>")

       Prhtml("<b>Followup question: </b> <br>")
       
       if ans4<=5 then 
          PrHtml("You chose " & newlist(ans4) & "<br>")
          Prhtml(NewCorrections(ans4))
        else
           Prhtml("You did not choose one of the conclusions.")
        end if
       
     prhtml("</td></tr></table>")

   
   
    function ChkAns(a,b)
      if abs(a-b) < 0.0015 then
        chkans=1
      else
        chkans=0
      end if
    end function
    %>
    <p align = "center"> Click <a href="zConfIntForMuQues.asp">here</a> for a new question </p>
</body>

</html>