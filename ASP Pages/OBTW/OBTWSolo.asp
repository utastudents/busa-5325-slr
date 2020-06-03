<html>
<head>

</head>


<body>
<form action=Main.asp method=POST enctype="application/x-www-form-urlencoded">
   
<% name = ""
      If session("Name") <> "" Then name = Replace(session("Name"), " ", "-") %>
<br>
1. Enter your name, last name first
    <br>
     <input name="name" value="">
     <br> <br>



2. Choose the number of opponents (Default is 4 Opponents)
    <br>
    <select name="NumOpp">
      <option value="4"> 4 Opponent
      <option value="3"> 3 Opponents
      <option value="2"> 2 Opponents
      <option value="1"> 1 Opponents
    </select>
    <% Session("NumOpp") ="4" %>

<br>
<br>

3. Choose the Difficulty of the Opponents <br>
   (Default is highest level of Difficulty) <br>
    <select name="Difficulty">
      <option value="4"> Students make 100 on exams
      <option value="3"> Students make B's or C's on exams
      <option value="2"> Students make D's on exams
      <option value="1"> Students failed every exam



    </select>
   <% Session("Difficulty") = "4" %>
   <% Session("GameStart") = "1" %>
<br> <br>


<input type="submit" value="Please click here to continue" name="END">

<p><!--#include file = "first.inc"--> </p>

</body>

</html>

