<html>
<head>

</head>


<body>

<form method="Post" action="Main.asp" >
<br> Enter only the ID numbers, the program <br> 
will look up the sizes <br> <br>

<%
   
   dim strRowlabels()
   dim SampleValues()   
   dim n 
   dim Average
   dim NumberSelectedIDs
   n = 10
   
    redim strRowlabels(n)
    redim SampleValues(n) 
   'valuelist = session("SelectedIds")   
   
   'SampleValues = split(valuelist,",") 

   'prhtml(valuelist & "<br>")
  ' prhtml(" Size of samplevalues?"  & ubound(SampleValues) & " <br>")
  
   'prhtml("Start variable = " & session("start"))
  

  Boxsizes = array(99, 1, 1, 1, 1, 1, 9, 1, 1, 1, 4, 5, 12, 1, 4, 8, 10, 16, 1, 6, 5, 4, 9, 24, 10, 4, 1, 12, 5, 16, 1, 4, 1, 1, 4, 12, 14, 4, 6, 4, 10, 1, 1, 7, 5, 8, 4, 9, 10, 27, 1, 1, 1, 5, 2, 12, 30, 1, 3, 6, 4, 1, 1, 28, 6, 1, 8, 1, 9, 10, 18, 4, 12, 6, 5, 3, 21, 5, 8, 2, 18, 1, 1, 24, 6, 30, 1, 1, 2, 16, 5, 1, 42, 3, 5, 1, 4,, 1, 3, 4, 8)

  


 call createtable()
 'call FindBoxSizes()
  Prhtml("<br> Once all the ID numbers are entered, press the button below <br> <br>")


 session("Start") = 1

'==============Subroutines follow =====================





sub CreateTable()
  dim stText 
  dim intIndex
  dim sum
  dim intNumberValues
  sum = 0
  intNumberValues = 0 
  prhtml("<br>")
  prhtml("<table  border =" & chr(34) & 1 & chr(34) & "   bgcolor = " & chr(34) & "Lavender" & chr(34) & ">")
  
  prhtml("<tr><td> </td><td> Enter Id Numbers </td><td> Box Size </td></td>") 
   for i = 1 to n
     
     prhtml("<tr>")
       Prhtml("<td>")
         prhtml("Box " & chr(64+i) & ":" )
       prhtml("</td> <td>")
         intindex = ""
         if session("Start") = 1 then
             intindex = request.form("box" & i)
         end if
         strtext = "<input     type="  & Chr(34) & "text"       & Chr(34) 
         strText = strtext & " name="  & Chr(34) & "box"  & (i) & Chr(34)
         strText = strtext & " Value=" & Chr(34) & intIndex     & chr(34)
         strtext = strtext & " width=" & chr(34) & "10"         & chr(34)
         strText = strtext & " ><br>"
         prhtml(strtext)
       prhtml("</td> <td>")
         strtext = ""
         'check first for a number and if a number
         if session("Start") = 1 and isnumeric(intindex) then
           'check for empty value
           if intindex <> "" then
              'check for value < 1 or > 100
              intindex = round(intindex,0)
              if (intindex < 1) or (intindex>100) then 
                 strtext = "Bad box ID"
              else
                 strtext = boxsizes(intindex)
                 SampleValues(i) = strText
                 sum = sum + cint(strText)
                 intNumberValues = intNumberValues +1
              end if
             
           end if
           prhtml(strtext) 
         end if
         prhtml(" </td>")
      prhtml("</tr>")
   next
   strText1 = ""
   strText2 = ""
   
   if (session("start") = 1) and (intNumberValues > 0) then 
      
       if (intNumberValues < n) then
         strText1 = "<p> You need " & n & " Ids. </p>"
         strText1 = strText1 & " You only have " & intNumberValues & "."
       else
         strText1 = "The average is "
         strText2 = round(sum/intNumberValues,4)
       end if
   end if
  prhtml("<tr><td> </td><td> " & strtext1 & " </td><td> " &  strtext2 & "</td></td>")
  prhtml("</table>")
end sub
   

sub prhtml(strText)
   response.write(strText)
end sub

 

%>

   <input type="submit" name="Submit" value="Average" />
     
  </form>
</body>

</html>

