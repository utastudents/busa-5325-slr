<head>

</head>
<body>

  <%
   Sub showform
    dim ComputerID 
    ComputerID = session("ComputerID")
  %>
  <form method="post">
      Choose an Assumption to Confirm <br>
      <table >
        <tr >
          <td >
 	
           <input type="radio" name="ANS" value="1"> Linearity <br>
           <input type="radio" name="ANS" value="2"> Indepedence <br>
           <input type="radio" name="ANS" value="3"> Normality <br>
           <input type="radio" name="ANS" value="4"> Equal Variance <br>
           <%  'New: Determine if CID has ever been entered 
             if len(trim(request.form("CID")))=0 and len(trim(ComputerID))) = 0 then
                ' if never entered from form and is not in computerid then request it
                response.write( " 2. Computer number &nbsp    <input type = " & chr(34) & "text"  & chr(34) & " name="& chr(34) & "cid" & chr(34) & "> <br>")
             else
                'if request value is available use it and store answer in computerid
                len(trim(ComputerID))) = 0 then ComputerID = request.form("CID")
                Session("ComputerID") = computerID
             end if
           %>
         </td>
       </tr>
      </table>
      <input type="submit" name="Submit" value="Submit" />
   </form>
  <%
    End Sub

    If Request.Form("Submit") = "" Then
       call showform()
    Else
       If Len(Trim(Request.Form("cid")))>0 Then
          call showform()
          Response.Write "<br> You choose to confirm Assumption " & Trim(Request.Form("ANS") & "<br> Computer ID is " & computerid)
          'Response.Write "<br />Thanks!!"
       Else
          Response.Write "You forgot to enter your computer number"
          call showform()
      End If
    End If

  %>
</body>