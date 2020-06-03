<html>
<head>

</head>

<body>
<form action=temp.asp method=POST enctype="application/x-www-form-urlencoded">
<h1 align="left"> Example of Basic Probabilililty Rules</h1>
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
   ATotalRow = array(0,0,0,0,0)
   ATotalRow(1) = int(rnd()*(grandtotal-3)+.5)+1
   ATotalRow(2) = int(rnd()*(grandtotal-atotalrow(1)-2)+.5)+1
   ATotalRow(3) = int(rnd()*(grandtotal-atotalrow(1)-atotalrow(2)-1)+.5)+1
   ATotalRow(4) = grandtotal-atotalrow(1)-atotalrow(2)-atotalrow(3)
   BTotalCol = array(0,0,0,0)

   
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

      
      
   

   RowDef="<td align=" & chr(34) & "right" & chr(34)  & "> "
'Creating HTML Table
   PrHtml("<table  align=" & chr(34) & "left" & chr(34) & "border=" & chr(34) & 1 & chr(34) & ">  ")

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

'Adding spaces under table
    for i=1 to 10
      PrHtml("<br>")
    next 

   'PrHtml( "ColChoice=" & colchoice & "<br>")
   'PrHtml("Rowchoice=" & Rowchoice & "<br>")
 
    PrHtml("A. Pr(" & ColumnLabel(colchoice) & ") = " & ColChoiceCount & " / " & _
             grandtotal & " = " & round(colchoicecount/grandtotal,3) & "<br>")
 
    PrHtml("B. Pr(" & RowLabel(Rowchoice) & ") = " & RowChoiceCount & " / " & _
             grandtotal & " = " & round(Rowchoicecount/grandtotal,3) & "<br>")

    PrHtml("C. Pr(not " & ColumnLabel(comchoice) & ") = 1 - (" & ComChoiceCount & _
             " / " & grandtotal & ") = " & round(1-Comchoicecount/grandtotal,3) & "<br>")


    PrHtml("D. Pr(" & columnlabel(colintchoice) & " and " & RowLabel(rowintchoice) & ") = " & _
              IntChoiceCount & " / " & grandtotal & " = " & round(intchoicecount/grandtotal,3) & "<br>")  

    PrHtml("E. Pr(" & columnlabel(colUnionchoice) & " or " & RowLabel(rowUnionchoice) & ") =" & _
             "Pr(" & columnlabel(colUnionchoice) & ") + Pr(" & rowlabel(rowunionchoice) & _  
              ") - Pr(" & columnlabel(colUnionchoice) & " and " & rowlabel(rowunionchoice) & _ 
              ") <br>")
        skipspaces(5)
        prhtml("= (" & Atotalrow(colUnionchoice) & " / " & grandtotal & ") + (" & _
              BtotalCol(rowunionchoice) & " / " & grandtotal & ") - (" & _
             contable(rowUnionchoice,colUnionchoice) & " / " & grandtotal & ")" & _ 
              " = " & round(UnionChoiceCount/grandtotal,3) & "<br>") 

    PrHtml("F. Pr(" & ColumnLabel(Colagivenbchoice) & " given " & Rowlabel(RowAgivenBchoice)& ") = " & _
             "Pr(" & columnlabel(Colagivenbchoice) & " and " & rowlabel(RowAgivenBchoice) & _  
              ") / Pr(" & rowlabel(RowAgivenBchoice) & _ 
              ") <br>")             

        skipspaces(5)
        prhtml("= " & contable(RowAgivenBchoice,Colagivenbchoice) & " / " & _
                BtotalCol(RowAgivenBchoice)  & " = " & _
                round(contable(RowAgivenBchoice,Colagivenbchoice)/BtotalCol(RowAgivenBchoice),3) & _
                "<br>") 

    PrHtml("F. Pr(" & Rowlabel(RowBgivenAchoice) & " given " & ColumnLabel(ColBgivenAchoice) & ") = " & _
             "Pr(" & columnlabel(ColBgivenAchoice) & " and " & rowlabel(RowBgivenAchoice) & _  
              ") / Pr(" & columnlabel(ColBgivenAchoice) & _ 
              ") <br>")             

        skipspaces(5)
        prhtml("= " & contable(RowBgivenAchoice,ColBgivenAchoice) & " / " & _
                AtotalRow(ColBgivenAchoice)  & " = " & _
                round(contable(RowBgivenAchoice,ColBgivenAchoice)/AtotalRow(ColBgivenAchoice),3) & _
                "<br>") 


  sub PrHtml(strHtml)
     response.write(strhtml)
  end sub

  sub SkipSpaces(IntSp)
    for intS = 1 to Intsp
      response.write("&nbsp;")
    next
  end sub


%> 
<br>
	
Click <a href="<%= request.servervariables("script_name") %>">here</a> for a new example

</body>

</html>