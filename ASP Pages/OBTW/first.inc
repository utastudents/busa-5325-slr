<%
 'There are at least three types of cards
 'Question cards, Technique cards and OBTW cards
 'Question cards have the text of the question and the appropriate solution
 'Technique cards have the name of the procedure and a list of the required assumptions
 'OBTW cards have ( the text of the card and the effect of the card: The effect consists of
'       (a) what assumptions are being affected, 
'	(b) the player being affected (You, One Opponent or All players), 
'	(c) whether it is a positive or negative effect,  
'       (d) more could be added later
 
    dim NumQues
    dim NumTech	
    dim NumReq
    dim NumCards	
    dim NumEffects
    Dim intInitialize
	
    dim strQuestions()	
    dim intSolutions()	
    dim strTechNames()
    dim strAssmList()
    dim strReqList()	
    dim strCardText()	
    dim strEffectList()
	
    NumQues=24	
    NumTech = 19
    NumReq=5  
    NumCards = 19
    intInitialize = 0	
	
    Redim strQuestions(NumQues)
    redim intSolutions(NumQues)
    redim strTechNames(NumTech)
    redim strReqList(NumTech,NumReq)
    redim strCardText(numcards)
    redim strEffectList(NumCards)
	
    Call LoadQuestions
    Call LoadRequirements	
    call LoadCards	
	
	
    Session("NumQues")  = NumQues	
    Session("NumTech")  = NumTech 
    Session("NumReg")   = NumReg 
    Session("NumCards") = NumCards
    Session("Initialize") = intInitialize
	
    Session("Questions")   = strQuestions
    Session("Solutions")   = intSolutions
    Session("Techniques")  = strTechNames
    Session("ReqList")     = strReqList
    Session("CardText")    = strCardText
    Session("EffectList")  = strEffectList
   

    response.write("<br> <br> " & intInitialize  )

   
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

end sub

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
     ' Response.write("<br> " & strTemp2 & "  " & strTechNames(i))
   next

end sub

Sub LoadCards()
  Dim Cards_n_Effects()
   
   Dim i
   Dim j
   Dim strTemp
   Dim strTemp2
   Dim strTemp3
   redim cards_n_Effects(NumCards)

     Cards_n_Effects(1)="(1122)Oh-By-The-Way: The data is non-normally distributed. Effects: All players must have a Small Sample size and a Moderate Sample size box checked, then they can choose Normality. If any Normality box is checked, it must be unchecked. Next Player's turn."
     Cards_n_Effects(2)="(1122)Oh-By-The-Way: The data is severely non-normally distributed.  Effects: : All players must have a Small Sample size and a Moderate Sample, and Large Sample size box checked, then they can choose Normality. If any Normality box is checked, it must be unchecked. Next Player's turn."
     Cards_n_Effects(3)="(1112)Oh-By-The-Way: The company remembers that they wish to control for the effect of an additional quantitative variable.  Effects: All players must be using a technique that allows multiple independent variables. Any Linearity box checked must be unchecked. Next Player's turn. "
     Cards_n_Effects(5)="(1132)Oh-By-The-Way:  The assumption of equal variance is violated. Effects: All Equal Variance boxes must be unchecked. Next Player's turn. "
     Cards_n_Effects(6)="(1112)Oh-By-The-Way:  The expected value of Y is not the specified linear function of the Xis.  Effects: All Linearity boxes must be unchecked. Next Player's turn. "
     Cards_n_Effects(7)="(2362)Oh-By-The-Way:  Data was artificially created by experimenter. Effect: An  opponent on your choice must uncheck their Sample boxes and Assumption boxes. Go again. "
     Cards_n_Effects(8)="(3362)Oh-By-The-Way: Bad data - Too many missing observations. Effect: You must uncheck all your boxes. Next Player's turn. "
     Cards_n_Effects(9)="(3362)Oh-By-The-Way: The instructor of your statistics course has a fake diploma. Effect: You must uncheck all your boxes. Next Player's turn."
     Cards_n_Effects(10)="(3362)Oh-By-The-Way: A software virus destroys your data. Effect: You must uncheck all your boxes. Next Player's turn. "
     Cards_n_Effects(11)="(2122)Oh-By-The-Way: An opponent believes that the data is non-normally distributed. Effects: An opponent must have a Small Sample size and a Moderate Sample size box checked, then they can choose Normality. If his/her Normality box is checked, it must be unchecked. Go again."
     Cards_n_Effects(12)="(2322)Oh-By-The-Way: An Opponent believes that the data is severely non-normally distributed.  Effects: An opponent must have a Small Sample size and a Moderate Sample, and Large Sample size box checked, then they can choose Normality. If her/his Normality box is checked, it must be unchecked. Go again."
     Cards_n_Effects(13)="(2132)Oh-By-The-Way:  An opponent believes that the assumption of equal variance is violated. Effects: Her/his  Equal Variance boxes must be unchecked. Go again. "
     Cards_n_Effects(14)="(2112)Oh-By-The-Way: An opponent believes that the expected value of Y is not the specified linear function of the Xis.  Effects: His/her Linearity boxes must be unchecked. Go again. "
     Cards_n_Effects(15)="(3251)Oh-By-The-Way: Your statistics professor comes by to visit. Effect: You get to check two of your boxes. Next player's turn."
     Cards_n_Effects(16)="(3351)Oh-By-The-Way: You spy on an opponent. Effect: You may check all the boxes that an opponent has checked. Next Player's turn."
     Cards_n_Effects(17)="(3171)Oh-By-The-Way: You buy an expert system for statistics. Effect: Starting your next turn you can choose up to two boxes per turn. Next Player's turn."
     Cards_n_Effects(18)="(3181)Oh-By-The-Way: You hire a statistical consultant.  Effect: Starting next turn you may choose up to two boxes per turn. Next Player's turn."
     Cards_n_Effects(19)="(2182)Oh-By-The-Way: An opponent's consultant quits. Effect: If an opponent has a consultant, then they lose that help. Next Player's turn."
     Cards_n_Effects(4)="(2172)Oh-By-The-Way: An opponent's expert system crashes. Effect: If an opponent has an expert system, then they lose that help. Go Again."
 
  For i = 1 To Numcards
     strCardText(i) = Right(    cards_n_Effects(i),    Len(  cards_n_Effects(i)  ) - 6)
     strEffectList(i) = mid(cards_n_Effects(i),2,4) 
   next
 
 end sub
 
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