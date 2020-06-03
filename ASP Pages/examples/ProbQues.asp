<html>
<head>

</head>


<body>
<form action=ProbSol.asp method=POST enctype="application/x-www-form-urlencoded">

<h1 align="center"> Basic Probability Rules Exercise</h1>
  

   <table align="center" width="600">

	
   <tr>
     <td>
      Enter your name, last name first
		<input style="width:350px" name="name" value=<%=Name%>>
      	        
     </td>
    </tr>
   <tr>
   </table><br>

<br>
 <%   
   
  dim ContTable()
   redim ConTable(3,4)

  
'Variable Definitions
   ColumnLabel = Array(" " ,"A1","A2","A3","A4")
   RowLabel = Array(" ","B1","B2","B3")

   randomize
   Grandtotal = int(rnd()*300+.5)+200
   
   'PrHtml("grandtotal =" & grandtotal)
   'Initialize Col totals
   ATotalRow = array(0,0,0,0,0)
   ATotalRow(1) = int(rnd()*(grandtotal-3)+.5)+1
   ATotalRow(2) = int(rnd()*(grandtotal-atotalrow(1)-2)+.5)+1
   ATotalRow(3) = int(rnd()*(grandtotal-atotalrow(1)-atotalrow(2)-1)+.5)+1
   ATotalRow(4) = grandtotal-atotalrow(1)-atotalrow(2)-atotalrow(3)
   BTotalCol = array(0,0,0,0)
   NumAttempts=0
   
   for jCol = 1 to 4
     
     contable(1,jcol) = int(rnd()*atotalrow(jcol)+.5)
     BtotalCol(1)=BtotalCol(1)+contable(1,jcol)   
     
     contable(2,jcol) = int(rnd()*(atotalrow(jcol)-contable(1,jcol))+.5)
     BtotalCol(2)=BtotalCol(2)+contable(2,jcol)   

     contable(3,jcol) = (atotalrow(jcol)-contable(1,jcol)-contable(2,jcol))
     BtotalCol(3)=BtotalCol(3)+contable(3,jcol)   
     
  
   next
 

'Determine what questions are being asked
   'A. Column choice for Column Prob
      ColChoice = round(rnd()*4+.5)
      ColChoiceCount = AtotalRow(colchoice)
   'B. Row choice for Row Prob
      RowChoice = round(rnd()*3+.5)
      RowChoiceCount = BtotalCol(Rowchoice)
   'C. Complementory choice will always be a Column Complement
      ComChoice = round(rnd()*4+.5)
      CoMChoiceCount = AtotalRow(comchoice)
   'D. Intersections
      RowIntChoice = round(rnd()*3+.5)
      ColIntChoice = round(rnd()*4+.5)
      IntChoiceCount = contable(RowIntChoice,ColIntChoice)
    'E Union
      RowUnionChoice = round(rnd()*3+.5)
      ColUnionChoice = round(rnd()*4+.5)
      UnionChoiceCount = Atotalrow(colunionchoice) + BtotalCol(rowUnionchoice) - contable(rowUnionchoice,colUnionchoice)
    'F A given B
      RowAGivenBChoice = round(rnd()*3+.5)
      ColAGivenBChoice = round(rnd()*4+.5)
      AGivenBNum = contable(RowAGivenBChoice,ColAGivenBChoice)
      AGivenBDen = AtotalRow(colAgivenbchoice)
    'G. B Given A
      RowBGivenAChoice = round(rnd()*3+.5)
      ColBGivenAChoice = round(rnd()*4+.5)
      BGivenANum = contable(RowBGivenAChoice,ColBGivenAChoice)
      BGivenADen = BtotalCol(RowBGivenAChoice)


     Session("contable")=contable
    
     Session("ColChoice") = Colchoice

     Session("RowChoice") = RowChoice

     Session("ComChoice") = ComChoice

     Session("RowIntChoice") = RowIntChoice
     Session("ColIntChoice") = ColIntChoice

     Session("RowUnionChoice")=RowUnionChoice
     Session("ColUnionChoice")=ColUnionChoice

     Session("RowAGivenBChoice")=RowAGivenBChoice
     Session("ColAGivenBChoice")=ColAGivenBChoice

     Session("RowBGivenAChoice") = RowBGivenAChoice
     Session("ColBGivenAChoice") = ColBGivenAChoice    
     Session("NumAttempts") = numattempts	

      RowDef="<td align=" & chr(34) & "right" & chr(34)  & "> "

     Name=""
     if Session("Name")<>"" then name=replace(session("Name")," ","-")

    PrHtml("<table  align=" & chr(34) & "center" & chr(34) & "border=" & chr(34) & 1 & chr(34) & ">  ")

   'Headings
    PrHtml("<tr>")
    PrHtml("<td></td>")
        for jcol=1 to 4
      PrHtml(rowdef &" <b>&nbsp;A" & jcol & "</b></td>")
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
    PrHtml(rowdef & grandtotal & " </td> ")
    PrHtml(" </tr> </table>")
    
    PrHtml("<p align=" & chr(34) & "center" & chr(34) & ">Given the above table, answer the following<br>")
    PrHtml("rounding your answer to three decimal places</p>")

    Prhtml("<table align=" & chr(34) & "center" & chr(34) & "><tr><td>")

    
      PrHtml("A. Pr(" & ColumnLabel(colchoice) & ") = <input  name=" & chr(34) & "AnsQuesA" & chr(34) & ">" )
      prhtml("</td></tr><tr><td>")

      PrHtml("B. Pr(" & RowLabel(Rowchoice) & ") = <input  name=" & chr(34) & "AnsQuesB" & chr(34) & ">")
      prhtml("</td></tr><tr><td>")
    
    
      PrHtml("C. Pr(not " & ColumnLabel(comchoice) & ") =<input  name=" & chr(34) & "AnsQuesC" & chr(34) & ">")
      prhtml("</td></tr><tr><td>")


      PrHtml("D. Pr(" & columnlabel(colintchoice) & " and " & RowLabel(rowintchoice) & _
          ") =<input  name=" & chr(34) & "AnsQuesD" & chr(34) & ">")
      prhtml("</td></tr><tr><td>")


      PrHtml("E. Pr(" & columnlabel(colUnionchoice) & " or " & RowLabel(rowUnionchoice) & _
          ") = <input  name=" & chr(34) & "AnsQuesE" & chr(34) & ">")
      prhtml("</td></tr><tr><td>")


      PrHtml("F. Pr(" & ColumnLabel(Colagivenbchoice) & " given " & Rowlabel(RowAgivenBchoice) & _
           ") =<input  name=" & chr(34) & "AnsQuesF" & chr(34) & ">")
       prhtml("</td></tr><tr><td>")


      PrHtml("G. Pr(" & Rowlabel(RowBgivenAchoice) & " given " & ColumnLabel(ColBgivenAchoice) & _
          ") =<input  name=" & chr(34) & "AnsQuesG" & chr(34) & ">")
    

    PrHtml("</td></tr></table>")

    sub PrHtml(strHtml)
       response.write(strhtml)
    end sub

    sub SkipSpaces(IntSp)
      for intS = 1 to Intsp
        response.write("&nbsp;")
      next
    end sub

 %><!--  -->

 


<p align="center">
  <input type="submit" value="Please click here to check answers" name="END">
</p>

<p align="center">
   Click <a align="center" href="<%= request.servervariables("script_name") %>">here</a> for a new exercise
</p>
</body>

</html>