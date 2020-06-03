<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>
<h1 align="center"> Basic Probability Rules Exercise </h1>
<br>
 <%   
   AnsA=request.form("AnsQuesA")
   AnsB=request.form("AnsQuesB")
   AnsC=request.form("AnsQuesC")
   AnsD=request.form("AnsQuesD")
   AnsE=request.form("AnsQuesE")
   AnsF=request.form("AnsQuesF")
   AnsG=request.form("AnsQuesG")
   Contable=SESSION("contable")
   Colchoice=  Session("ColChoice") 
   RowChoice=  Session("RowChoice") 
   ComChoice=  Session("ComChoice") 
   RowIntChoice=  Session("RowIntChoice") 
   ColIntChoice=  Session("ColIntChoice") 
   RowUnionChoice=  Session("RowUnionChoice")
   ColUnionChoice=  Session("ColUnionChoice")
   RowAGivenBChoice=  Session("RowAGivenBChoice")
   ColAGivenBChoice=  Session("ColAGivenBChoice")
   RowBGivenAChoice=  Session("RowBGivenAChoice") 
   ColBGivenAChoice =  Session("ColBGivenAChoice")
   NumAttempts = session("NumAttempts")
   NumAttempts = numattempts+1
   session("Numattempts") = numattempts

   ATotalRow = array(0,0,0,0,0)
   BTotalCol = array(0,0,0,0)
   ColumnLabel = Array(" " ,"A1","A2","A3","A4")
   RowLabel = Array(" ","B1","B2","B3")
   grandtotal=0
  
   For irow = 1 to 3   
      for jCol = 1 to 4
        BtotalCol(irow)=BtotalCol(irow)+contable(irow,jcol) 
      next     
   next
   For jcol = 1 to 4   
      for irow = 1 to 3
        AtotalRow(jcol)=AtotalRow(jcol)+contable(irow,jcol) 
        grandtotal=grandtotal+contable(irow,jcol)
     next     
   next

       
   
      RowDef="<td align=" & chr(34) & "right" & chr(34)  & "> "
    PrHtml("<table  align=" & chr(34) & "center" & chr(34) & "border=" & chr(34) & 1 & chr(34) & ">  ")

  'Determine the answers to the questions being asked
   'A. Column choice for Column Prob
        ColChoiceCount = AtotalRow(colchoice)
        QuesASol = round(colchoicecount/grandtotal,3)

   'B. Row choice for Row Prob
        RowChoiceCount = BtotalCol(Rowchoice)
        QuesBSol = round(RowChoiceCount/grandtotal,3)

   'C. Complementory choice will always be a Column Complement
        CoMChoiceCount = AtotalRow(comchoice)
        QuesCSol = round((1-CoMChoiceCount/grandtotal),3)

   'D. Intersections
        IntChoiceCount = contable(RowIntChoice,ColIntChoice)
        QuesDSol = round(IntChoiceCount/grandtotal,3)      

    'E Union
        UnionChoiceCount = Atotalrow(colunionchoice) + BtotalCol(rowUnionchoice) - contable(rowUnionchoice,colUnionchoice)
         QuesESol = round(UnionChoiceCount/grandtotal,3)      

    'F A given B
      AGivenBNum = contable(RowAGivenBChoice,ColAGivenBChoice)
      AGivenBDen = btotalcol(rowAgivenbchoice)
      QuesFSol = round(AGivenBNum/AGivenBDen,3)    

    'G. B Given A
      BGivenANum = contable(RowBGivenAChoice,ColBGivenAChoice)
      BGivenADen = Atotalrow(colBGivenAChoice)
      
      QuesGSol = round(BGivenANum/BGivenADen,3) 
'Headings
    prhtml("<p align=" & chr(34) & "center" & chr(34) & "> Number of attempts = " & numattempts & "<p>")
    PrHtml("<tr>")
    PrHtml("<td></td>")
        for jcol=1 to 4
          PrHtml(rowdef & " <b>&nbsp;A" & jcol & "</b></td>")
       next
    PrHtml(rowdef & "<b>Total</b></td></tr>")
   
   
    for irow= 1 to 3
      PrHtml("<tr>")
      PrHtml(rowdef &" <b>" & Rowlabel(irow) & "</b></td>")
       for jcol = 1 to 4
         PrHtml(rowdef & contable(irow,jcol) & " </td> ")
       next
       PrHtml(rowdef & BtotalCol(irow) & " </td> </tr>")
    next
   
  
    PrHtml("<tr>")
    PrHtml(rowdef &" <b>Total</b></td>")
     for i = 1 to 4
        PrHtml(rowdef & atotalrow(i) & " </td> ")
     next
    PrHtml(rowdef & grandtotal & " </td> </tr>")

    PrHtml(" </table>")

    Prhtml("<table align=" & chr(34) & "center" & chr(34) & "><tr><td>")

    prhtml("<br>In case of an error, if you are using Internet Explorer, then <br>" & _
              "go Back to the previous page and try again. The program will <br>" & _
              "keep track of the number of attempts.<br><br>")
 
       prhtml("</td></tr><tr><td>")  
   'Question A
    PrHtml("Question A: ") 
      if isnumeric(ansa) = true then
    
        If chkans(ansa,QuesASol) =1 then
         'If csng(ansa)=csng(QuesASol) then

           PrHtml("Correct: ")
           PrHtml("Pr(" & ColumnLabel(colchoice) & ") = " & ColChoiceCount & " / " & _
             grandtotal & " = " & round(colchoicecount/grandtotal,3) & "<br>")
         Else
            Prhtml("Incorrect: Your answer of " & ansa & " is wrong. <br>")
         end if
      elseif len(ansa) =0 then
         prhtml("You did not enter an answer to Question A. <br> <br>")
      else
            Prhtml("Your answer of " & ansa & " is not a number. Back up and try again.<br>")
      end if
       prhtml("</td></tr><tr><td>")
   'Question B 
      PrHtml("Question B: ") 
      if isnumeric(ansb) = true then
         If chkans(ansb,QuesbSol) =1 then
        
         
           PrHtml("Correct: ")
         PrHtml("Pr(" & RowLabel(Rowchoice) & ") = " & RowChoiceCount & " / " & _
             grandtotal & " = " & round(Rowchoicecount/grandtotal,3) & "<br>")
         Else
            Prhtml("Incorrect. Your answer of " & ansb & " is wrong. <br>")
         end if
      elseif len(ansb) =0 then
         prhtml("You did not enter an answer to Question B. <br> <br>")
      else
            Prhtml("Your answer of " & ansB & " is not a number. Back up and try again.<br>")
      end if
        prhtml("</td></tr><tr><td>")
  
 
      'Question c 
      PrHtml("Question C: ") 
      if isnumeric(ansc) = true then
        
         If chkans(ansc,QuescSol) =1 then
           PrHtml("Correct: ")
           PrHtml("Pr(not " & ColumnLabel(comchoice) & ") = 1 - (" & ComChoiceCount & _
             " / " & grandtotal & ") = " & round(1-Comchoicecount/grandtotal,3) & "<br>")
         Else
            Prhtml("Incorrect. Your answer of " & ansc & " is wrong. <br>")
         end if
      elseif len(ansc) =0 then
         prhtml("You did not enter an answer to Question C. <br> <br>")
      else
            Prhtml("Your answer of " & ansC & " is not a number. Back up and try again.<br>")
      end if
      prhtml("</td></tr><tr><td>")
   'Question D 
      PrHtml("Question D: ") 
      if isnumeric(ansD) = true then
        
        If chkans(ansd,QuesdSol) =1 then
           PrHtml("Correct: ")
           PrHtml("Pr(" & columnlabel(colintchoice) & " and " & RowLabel(rowintchoice) & ") = " & _
              IntChoiceCount & " / " & grandtotal & " = " & round(intchoicecount/grandtotal,3) & "<br>")  
         Else
            Prhtml("Incorrect. Your answer of " & ansD & " is wrong. <br>")
         end if
      elseif len(ansD) =0 then
         prhtml("You did not enter an answer to Question D. <br> <br>")
      else
            Prhtml("Your answer of " & ansD & " is not a number. Back up and try again.<br>")
      end if
      prhtml("</td></tr><tr><td>")
   'Question E 
      PrHtml("Question E: ") 
      if isnumeric(ansE) = true then
        
         If chkans(anse,QueseSol) =1 then
           PrHtml("Correct: ")
           PrHtml("Pr(" & columnlabel(colUnionchoice) & " or " & RowLabel(rowUnionchoice) & ") =" & _
             "Pr(" & columnlabel(colUnionchoice) & ") + Pr(" & rowlabel(rowunionchoice) & _  
              ") - Pr(" & columnlabel(colUnionchoice) & " and " & rowlabel(rowunionchoice) & _ 
              ") <br>")
            skipspaces(5)
            prhtml("= (" & Atotalrow(colUnionchoice) & " / " & grandtotal & ") + (" & _
              BtotalCol(rowunionchoice) & " / " & grandtotal & ") - (" & _
             contable(rowUnionchoice,colUnionchoice) & " / " & grandtotal & ")" & _ 
              " = " & round(UnionChoiceCount/grandtotal,3) & "<br>")   
         Else
            Prhtml("Incorrect. Your answer of " & ansE & " is wrong. <br>")
         end if
      elseif len(ansE) =0 then
         prhtml("You did not enter an answer to Question E. <br> <br>")
      else
            Prhtml("Your answer of " & ansE & " is not a number. Back up and try again.<br>")
      end if
      prhtml("</td></tr><tr><td>")    
   'Question f 
      PrHtml("Question F: ") 
      if isnumeric(ansF) = true then
        
        If chkans(ansf,QuesfSol) =1 then
           PrHtml("Correct: ")
           PrHtml("Pr(" & ColumnLabel(Colagivenbchoice) & " given " & Rowlabel(RowAgivenBchoice)& ") = " & _
             "Pr(" & columnlabel(Colagivenbchoice) & " and " & rowlabel(RowAgivenBchoice) & _  
              ") / Pr(" & rowlabel(RowAgivenBchoice) & _ 
              ") <br>")             

           skipspaces(5)
           prhtml("= (" & contable(RowAgivenBchoice,Colagivenbchoice) & "/" & grandtotal & ") / (" & _
                BtotalCol(RowAgivenBchoice) & "/" & grandtotal & ")" & " = " & _
                contable(RowAgivenBchoice,Colagivenbchoice) & " / " & BtotalCol(RowAgivenBchoice) & " = " & _
                round(contable(RowAgivenBchoice,Colagivenbchoice)/BtotalCol(RowAgivenBchoice),3) & _
              "<br>") 
         Else
            Prhtml("Incorrect. Your answer of " & ansF & " is wrong. <br>")
         end if
      elseif len(ansF) =0 then
         prhtml("You did not enter an answer to Question F. <br> <br>")
      else
            Prhtml("Your answer of " & ansF & " is not a number. Back up and try again.<br>")
      end if
      prhtml("</td></tr><tr><td>")
'Question G 
      PrHtml("Question G: ") 
      if isnumeric(ansG) = true then
        
        If chkans(ansg,QuesgSol) =1 then
           PrHtml("Correct: ")
           PrHtml("Pr(" & Rowlabel(RowBgivenAchoice) & " given " & ColumnLabel(ColBgivenAchoice) & ") = " & _
             "Pr(" & columnlabel(ColBgivenAchoice) & " and " & rowlabel(RowBgivenAchoice) & _  
              ") / Pr(" & Columnlabel(colBgivenAchoice) & _ 
              ") <br>")             

           skipspaces(5)
           prhtml("= (" & contable(RowBgivenAchoice,ColBgivenAchoice) & "/" & grandtotal & ") / (" & _
                AtotalRow(ColBgivenAchoice)  & "/" & grandtotal & ") = " & _
                contable(RowBgivenAchoice,ColBgivenAchoice) & " / " & _
                AtotalRow(ColBgivenAchoice)  & " = " & _
                round(contable(RowBgivenAchoice,ColBgivenAchoice)/AtotalRow(ColBgivenAchoice),3) & _
                "<br>") 
         Else
            Prhtml("Incorrect. Your answer of " & ansG & " is wrong. <br>")
         end if
      elseif len(ansG) =0 then
         prhtml("You did not enter an answer to Question G. <br> <br>")
      else
            Prhtml("Your answer of " & ansG & " is not a number. Back up and try again.<br>")
      end if
      

     prhtml("</td></tr></table>")
    sub PrHtml(strHtml)
       response.write(strhtml)
    end sub

    sub SkipSpaces(IntSp)
      for intS = 1 to Intsp
        response.write("&nbsp;")
      next
    end sub
    function ChkAns(a,b)
      if abs(a-b) < 0.0015 then
        chkans=1
      else
        chkans=0
      end if
    end function

   response.write("<br>")



   
  %>

Click <a href="Probques.asp">here</a> for a new question
</body>

</html>