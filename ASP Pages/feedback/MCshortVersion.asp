
<html>

<head>



<title>My Online Survey</title>


</head>

<body >

<form action=secondMCmulti.asp method=POST enctype="application/x-www-form-urlencoded">
   Please choose your answer, enter your computer number and click the Send button <br>

   <% 'Note that both the answer and the computer number will be checked for correct length in the secondMCmulti.asp file %>

   <table >
     <tr >
       <td >
	
    
        <input type="radio" name="ANS" value="1"> A <br>
              

        <input type="radio" name="ANS" value="2" > B <br>


        <input type="radio" name="ANS" value="3"> C <br>
              

        <input type="radio" name="ANS" value="4" > D <br>
              

        <input type="radio" name="ANS" value="5" > E <br> 

     


	2. Computer number &nbsp    <input type = "text" name="cid" > <br>
      
     </td>
   </tr>
 </table>


												


<input type="submit" value="Send" name="END"></p>


</form>







</body>

</html>