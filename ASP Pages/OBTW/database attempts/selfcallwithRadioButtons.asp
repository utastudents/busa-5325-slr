<head>

</head>
<body>

  <%
   Sub showform
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
           
          2. Computer number &nbsp    <input type = "text" name="cid" > <br>
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
          Response.Write "<br> You choose to confirm Assumption " & Trim(Request.Form("ANS"))
          'Response.Write "<br />Thanks!!"
       Else
          Response.Write "You forgot to enter your computer number"
          call showform()
      End If
    End If

  %>
</body>