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
    <select name="NumOpp">
      <option value="4"> Students make 100 on exams
      <option value="3"> Students make B's or C's on exams
      <option value="2"> Students make D's on exams
      <option value="1"> Students failed every exam



    </select>
   <% Session("Difficulty") = "4" %>
   <% Session("GameStart") = "1" %>
<br> <br>


<input type="submit" value="Please click here to continue" name="END">

<%
  'Total Number of Questions Available   
    dim NumQues
    dim strQuestions()
    NumQues=24
    Session("NumQues") = NumQues
    Redim strQuestions(NumQues)
   
  'Techniques that Go with Each Question
    dim intSolutions()
    redim intSolutions(NumQues)

  'Load and Store Questions and Solutions 
    Call LoadQuestions
     'Session("Questions") =strQuestions
     'Session("Solutions") = intSolutions

  'Total Number of Techniques Available
    dim NumTech
    NumTech = 19
    Session("NumTech") = NumTech
    dim strTechNames()
    dim strAssmList()
    redim strTechNames(NumTech)

  'Total Number of Requirements 4 Assumptions (L, I, N, E) plus whether Pop can be normal or not 
    dim NumReq
    dim strReqList()
    NumReq=5    
    redim strReqList(NumTech,NumReq)

  'Load and Store Techniques and their Requirements 
     Call LoadRequirements
     'Session("Techniques") =strTechNames
     'Session("ReqList") = strReqList

Sub LoadQuestions()
   Dim Techs_n_Ques()
   
   Dim i 
   
   
   ReDim Techs_n_Ques(NumQues)
 
  'First Load the Questions
    Techs_n_Ques(1) = "(18)Does the increase in the mean salary with each year of experience change with the number of employees in the company?"
    Techs_n_Ques(2) = "( 6)You wish to estimate all possible diffences in  average executives salaries across four types of companies: consumer products, utilities, industrial-high tech, and financial services."
    Techs_n_Ques(3) = "( 5)Does the IRS assess the same aveage amount of fines for Manufacturing as opposed to Service Industries?"
    Techs_n_Ques(4) = "( 9)What will be  the average final exam grade for all students how make a 75 on exam 1 and 92 on exam 2?"
    Techs_n_Ques(5) = "( 3)Knowing only sample data what is the average IRS penalties?"
    Techs_n_Ques(6) = "(16)After determining that there is no interaction of machine and manufactuer,  is there a difference in average repair cost  among four types of machines?"
    Techs_n_Ques(7) = "(10)What will be the salary of a person with a given number of years experience?"
    Techs_n_Ques(8) = "(17)There is interaction of machine and manufactuer,  what is differences in average repair cost  between combinations of machine and manufacturers?"
    Techs_n_Ques(9) = "( 6)What will be average difference in salaries between the positions of assistant vice-president and supervisor?"
    Techs_n_Ques(10) = "(13)Among blocking by position (CEOs, CFOs, and CIOs), is there any difference in the average salaries between men and women?"
    Techs_n_Ques(11) = "( 8)After adjusting for the GMAT score, what is the change in  the average of the final grades for for each hour studied?"
    Techs_n_Ques(12) = "( 5)Is there any difference in the average of executive salaries among the different ranks (vice-president, assistant vice-president, or supervisor)?"
    Techs_n_Ques(13) = "( 7)Is undergraduate major related to the letter grade in course?"
    Techs_n_Ques(14) = "(15)People were  sampled from  combinations of gender and city. You wish to know if the effect of gender on salary depends on city."
    Techs_n_Ques(15) = "(12)For all students with a specific GPA, is there any difference in the average final grades between two majors?"""
    Techs_n_Ques(16) = "(14)After blocking by the number of statistics classes (0, 1, or 2), what is the difference in average grade among the  majors ?"
    Techs_n_Ques(17) = "(11)Can years experience and/or years of education  help predict an executive's salary?."
    Techs_n_Ques(18) = "(12)Is there evidence of discrimination (do male executives get paid more than women executives with the same experience)?"
    Techs_n_Ques(19) = "( 4)Based only on a sample will the average of the final grades for all students who could have taken the class be larger than 70?"
    Techs_n_Ques(20) = "( 9)A delivery service wishes to determine the average delivery costs for all deliveries made 20 miles from the store."
    Techs_n_Ques(21) = "( 2)A student wishes to know what percent of the time MBA students will use statistics at work."
    Techs_n_Ques(22) = "( 1)Given you only know the variation in salaries of all vice-presidents, based on a sample what is the  average salaries of all vice-presidents?"
    Techs_n_Ques(23) = "( 7)Does a person's opinion about capital punishment depend on their party affiliation?"
    Techs_n_Ques(24) = "(19)After determining that there is no interaction of machine and manufactuer,  what is the differences in average repair cost  among four types of machines?"


  'randomize the question order
    Call RandomSort(Techs_n_Ques)
   
   'seperate the questions and solution techniqes
  For i = 1 To NumQues
     strQuestions(i) = Right(Techs_n_Ques(i), Len(Techs_n_Ques(i)) - 4)
     intsolutions(i) = Trim(Mid(Techs_n_Ques(i), 2, 2))
  '   response.write("<br>")
  '   response.write(  intSolutions(i) & ", " & strQuestions(i)  )
  Next 
     
End Sub

Sub LoadRequirements()

 Dim Assms_n_Techs()
   
 Dim i
 Dim j
 Dim strTemp
 Dim strTemp2
 Dim strTemp3
   
  
   ReDim Assms_n_Techs(NumTech)
    Assms_n_Techs(1) = "(01100) Z conf. interval (CI) for a  mean"
    Assms_n_Techs(2) = "(10010) Z CI for a  proportion"
    Assms_n_Techs(3) = "(01101) t CI for a  mean"
    Assms_n_Techs(4) = "(01101) t test for a  mean"
    Assms_n_Techs(5) = "(01111) One Way ANOVA F test"
    Assms_n_Techs(6) = "(01111) One-Way ANOVA Tukey "
    Assms_n_Techs(7) = "(10010) Chi-Square  contingency table"
    Assms_n_Techs(8) = "(11111) CI for a Slope"
    Assms_n_Techs(9) = "(11111) CI for a mean given Xi values"
    Assms_n_Techs(10) = "(11111) Prediction interval for Y | Xi values"
    Assms_n_Techs(11) = "(11111) Regression F test of Model"
    Assms_n_Techs(12) = "(11111) Dummy Variable Regression"
    Assms_n_Techs(13) = "(01111) Randomized Block F test"
    Assms_n_Techs(14) = "(01111) Random Block Tukey Procedure"
    Assms_n_Techs(15) = "(01111) 2-Way ANOVA F test - Interaction"
    Assms_n_Techs(16) = "(01111) 2-Way ANOVA F test - Factor Means"
    Assms_n_Techs(17) = "(01111) 2-Way ANOVA Tukey - Treatment Means"
    Assms_n_Techs(18) = "(11111) Test of a Slope "
    Assms_n_Techs(19) = "(01111) 2 Way ANOVA Tukey Factor Means"
   'seperate the questions and solution techniqes
  
  For i = 1 To NumTech
     strTechNames(i) = Right(Assms_n_Techs(i), Len(Assms_n_Techs(i)) - (NumReq + 3))
     strTemp = Trim(Mid(Assms_n_Techs(i), 2, 5))
     strTemp2 = ""
     For j = 1 To NumReq
        strTemp3 = Mid(strTemp, j, 1)
        strReqList(i, j) = strTemp3
        strTemp2 = strTemp2 & strReqList(i, j)
     Next
     Response.write(&"<br> " & strTemp2 & "  " & strTechNames(i))

  Next

End Sub

Sub RandomSort(arr)
  Dim strTemp  
  Dim i 
  Dim j 

  Dim lngMax 
  Dim sngRand()
  Dim sngTempRand
  sngTempRand = 2.01
  lngMax = UBound(arr)
  ReDim sngRand(lngMax)
  For i = 1 To lngMax
     randomize
     sngRand(i) = Rnd()
  Next 

  For i = 1 To lngMax - 1
    For j = i + 1 To lngMax
      If sngTempRand > sngRand(j) Then
        sngTempRand = sngRand(i)
        sngRand(i) = sngRand(j)
        sngRand(j) = sngTempRand
 
        strTemp = arr(i)
        arr(i) = arr(j)
        arr(j) = strTemp
      End If
    Next 
  Next 
End Sub

 %>



</body>

</html>

