<html>
<head>

</head>


<body>

<form method="Post" action="Main.asp" >

<%

  'the values of these three variables will be submitted from the Start form
    dim strName
    Dim intNumOpp
    Dim intDifficulty

  'The values of these variables are stored as session variables and are scaler
    dim NumQues              'Number of Questions
    dim NumTech	             'Number of different Stat Techniques
    dim NumReq               'L. I. N. E. along with 3 sample sizes
    dim NumCards	     'Number of OBTW cards 
    dim NumEffects           'Number of Effects of OBTW cards: # of opp it effects, etc
    Dim intInitializing      '0 if this is the  initiating run or start, 1 if continuation of game
    Dim intHumanPlayerChoice 'this is an integer from the radio button indicating players game choice

  'The values of these variables are stored as session variables and are arrays	
    dim strQuestions	     'the text of the questions
    dim intSolutions         'the correct technique to solve each question
    dim strTechNames         'The technique names
    dim strAssmList          'The names of the assumptions or requirments
    dim strReqList	     '
    dim strCardText	     'the text on the OBTW card
    dim strEffectList        'the list of the effects of each OBTW in the form of a number
                               'such as 1352 where each i-th digit is the value of one effect
    dim strCardsPlayed       'Values to be displayed indicting what assumptions have been
                               'validated and the sample sizes                         

  'Scalers created in this form
    dim intQuesNum           'contains the question number
    dim intActPlayerNum      'contains the active player number
    dim intPreviousPlayerNum 'contains the number for the previous player                                 
    dim intTurnNumber        'contains the turn number
    dim sngProbOfOBTW        'contains the probability of getting an OBTW card
    dim strCurrentCardText   'contains the text on the current card
    dim intComputerPlayerChoice  'contains values 1-4 for the four assumptions                  

  'Arrays created in this form
    
    dim strRowNames          'This will have the names of requirements

  'Values came from the start from
     NumQues=Session("NumQues") 	
     NumTech= Session("NumTech") 
     NumReg= Session("NumReg") 
     NumCards=Session("NumCards") 
     intInitializing = Session("Initialize")
   
     strQuestions=Session("Questions") 
     intSolutions=Session("Solutions") 
     strTechNames=Session("Techniques")
     strReqList=Session("ReqList")     
     strCardText=Session("CardText")   
     strEffectList=Session("EffectList")

       'list of the effects are in comments
       'at the top of the subrountine that 
       'displays the players choices

 'Values assigned here
    sngProbOfOBTW = 0.00  'force the obtw cards never to appear

 if intInitializing = 0 then
     'Values from the Start form

      strName = Request.Form("Name")
      intNumOpp = Request.Form("numopp")
      intDifficulty = Request.Form("difficulty")
      
    'Save these as session variables
      Session("Name") = strName
      Session("NumOpp") = intNumOpp
      Session("Difficulty")=intDifficulty
      Session("Initialize") = 1
 

    'Variables created for the first time and needed every turn
      intQuesNum   = 1        'Incremented after a question is solved
      intActPlayerNum = 1     'Incremented when OBTW card says next player
      intTurnNumber = 1       'Incremented when submit button is clicked
	
    'Recall the Session variables from above and put them in local variables      
      strName = session("Name")
      intNumOpp = session("NumOpp")
      intdifficulty = session("Difficulty")  
  

    'Store these for use next time
      session("QuesNum") = intQuesNum
      session("ActPlayerNum")=intActPlayerNum
      session("TurnNumber") = intTurnNumber
      call initializeTableValues
      'response.write("TurnNumber is " & session("TurnNumber") & " <br>")  
     	  
else

   'the following commands are run only after the initial turn
     
     intNumOpp = Session("NumOpp")
     strName = Session("Name")
     intDifficulty = Session("Difficulty")
     intTurnNumber = session("TurnNumber")
     intPreviousPlayerNum = session("ActPlayerNum")                                                                  
     
   'The turn number will increment and go to the next card
   'except when the problem will be solved or when a player
   'is allowed two turns
     intTurnNumber = intTurnNumber+1
     Session("TurnNumber") = intTurnNumber

   'recall the active player number and question number
     intActPlayerNum = session("ActPlayerNum")
     intQuesNum = session("QuesNum")

       'Reset CardsPlayed session array based on the values in the radio buttons
       'from the previous turn
         call updateCardsPlayed
	 ' intHumanPlayerChoice = Request.Form("AssmChoice")
	 ' PrHTML("Player choose " & intHumanPlayerChoice & "<br>")

  end if

  'THE FOLLOWIN CODE IS RUN EACH TURN

  'Recall the cards on the table 
  'Notice that on the first turn no cards have played
  'Recall the name of the rows in the table presented

   strCardsPlayed = Session("CardsPlayed")  
   strRowNames = Session("RowNames")
  
  'call sub that make the computer player's choices
     if (intActPlayerNum > 1) then Call ComputerPlayer                                             
  
  Call CreateGameWindow
   prhtml("Question " & intQuesNum & "," & "Player " & intPreviousPlayerNum & ", TurnNumber " & intTurnNumber & " <br>")        

   
  'check the text of the current card and see if it says "Go again"
    'if it doesn't say Go again then increment the active player number
    if inStr(strCurrentCardText,"Go again")>0 then
        'do not not increment
    else
        intActPlayerNum = intActPlayerNum+1
    end if  

    'total number of players is the number of opponents plus 1
    if intActPlayerNum > (intNumOpp+1) then intActPlayerNum = 1                        'new
    session("ActPlayerNum")=intActPlayerNum
   


' ==============================================================
'
'                Subroutines start here
'
' ==============================================================

sub prhtml(strText)
   response.write(strText)
end sub

Sub initializeTableValues()
   'This sub creates the names of the rows for the table and puts dashes for the values
   Dim i 
   Dim j 
   Dim strCardsPlayedStart
   ReDim strCardsPlayedStart(7, intNumOpp + 1)
   strRowNamesStart = Array("Linearity", "Independence", "Normality", "Equal Variance","Small Sample", "Large Sample",  "Very Large Sample")
   For i = 1 To intNumOpp + 1
      For j = 1 To 7
        strCardsPlayedStart(j, i) = "-"
        if j=5 then strCardsPlayedStart(j,i) ="X"
      Next 
    Next 
    Session("RowNames") = strRowNamesStart
   Session("CardsPlayed") = strCardsPlayedStart
End Sub

Sub DisplayCurrentQuestion()

 prhtml("<h5> Question " & intQuesNum & " </h5> ")
 'prhtml("<table border= " & chr(34) & "2" & chr(34)& " >")
 '     prhtml("<tr> <td>")
         prhtml(strQuestions(intQuesNum))
 '     prhtml("</td> </tr> ")
 'prhtml("</table  <br>")

end sub

Sub DisplayCurrentCard()
  Randomize
  sngProbOfOBTW = 0.20
  dim strText1
  dim strText2
  dim sngRand
  sngRand = rnd()

 ' if sngRand > sngProbOfOBTW then
 if sngRand > 0.0 then
 
    strText1 = "An Assumption has been validated. Click an Assumption or a larger Sample check box."
    strText2 = "<br> Click End Turn"
  else
    strText1 =  strCardText(Randbetween_1_and(numcards)) 
    strCurrentCardText = strText1
    Session("CurrentCardText") = strText1

'=========================change next line depending on the effect of the card =====
 
    strText2 = "<br> Click End Turn"
  end if

 
  prhtml("<h5> Card Drawn at Turn " & intTurnNumber & " </h5> ")
 ' prhtml("<table border= " & chr(34) & "2" & chr(34)& " >")
 '     prhtml("<tr> <td>")
         prhtml(strText1 & strText2)
 '     prhtml("</td> </tr> ")
 ' prhtml("</table  <br>")

end sub

Function Randbetween_1_and(intMax )
   Randbetween_1_and = Int(Rnd() * intMax) + 1
End Function


sub PrintPlayerTableAndPLayerChoices()

  ' prhtml("<table border= " & chr(34) & "2" & chr(34)& " >")
  '    prhtml("<tr> <td>")
         Call PrintPlayerTable
       Prhtml("</td>  <td>")
         call DisplayAssumptionChoices
    
  '    prhtml("</td> </tr> ")
  ' prhtml("</table  <br>")

end sub

Sub PrintPlayerTable()

   prhtml("<h5> Score Card Before Turn " & intTurnNumber & " </h5> ") 
   'prhtml("# of opponents" & intNumOpp )
   PrHTML ("<Table border=" & Chr(34) & "1" & Chr(34) & " style=" & chr(34) & "margin-left:10px" & Chr(34) & ">")
    style="margin-left:10px"
   PrHTML ("<tr>")
        PrHTML ("<td>")
        PrHTML ("Requirement")
        PrHTML ("</td>")
   
    'Second column is player's name
  
        PrHTML ("<td>")
        PrHTML (strName)
        PrHTML ("</td>")
  
   'Next columns are the computer opponents
        For jcol = 1 To intNumOpp
             PrHTML ("<td>")
             PrHTML ("Player " & (jcol+1))
             PrHTML ("</td>")
        Next 
      PrHTML ("</tr>")
        
    'This prints the current values of the assumptions that have been checked    
       For irow = 1 To 6
           PrHTML ("<tr>")
           intCurrentCol = 0
           For jcol = 1 To intNumOpp + 2
              PrHTML ("<td>")
              If jcol = 1 Then
                  PrHTML (strRowNames(irow-1))
              Else
                  PrHTML (strCardsPlayed(irow, jcol - 1))
              End If
              PrHTML ("</td>")
            Next 
            PrHTML ("</tr>")
        Next 
        PrHTML ("</table>")
     prhtml("<br>")
End Sub

sub CreateGameWindow()
  PrHTML("<Table border=" & Chr(34) & "1" & Chr(34) & ">")
    PrHTML("<tr>")
        PrHTML ("<td>")
          Call DisplayCurrentQuestion
        PrHtml("</td>")
    'Prhtml("</tr>")
   
    'PrHTML("<tr>")
        PrHTML ("<td>")
          Call DisplayCurrentCard
        PrHtml("</td>")
    Prhtml("</tr>")
   
    PrHTML("<tr>")
        PrHTML("<td>")
          call PrintPlayerTableAndPLayerChoices
        PrHtml("</td>")
    Prhtml("</tr>")
  prhtml("</Table>")
end sub

sub DisplayAssumptionChoices
  if intActPlayerNum = 1 then 
    ' PrHTML("Player choices go here")
     Prhtml("Choose one of the following <br>")
    
     'PrHTML ("<form action=" & Chr(34) & Chr(34) & ">")
	   dim sTemp
	   sTemp = "AssmChoice"
	   
       call CreateRadioButton(sTemp, "1", "Linearity")
	   call CreateRadioButton(sTemp, "2", "Independence")
	   call CreateRadioButton(sTemp, "3", "Normality")
	   call CreateRadioButton(sTemp, "4", "Equal Variance")
	   Prhtml("<br>")
	   call CreateRadioButton(sTemp, "5", "Obtain a Larger Sample")
	   Prhtml("<br>")
	   call CreateRadioButton(sTemp, "6", "Attempt to Solve Question")
	 'PrHTML ("</Select>")
  Else
     'if intPreviousPlayerNum = 1 and intTurnNumber > 1 then prHTML("Human player made choice " & intHumanPlayerChoice)   
     intComputerPlayerChoice = Session("ComputerPlayerChoise")                                                                 'new 
                    
     PrHTML("Player " & intActPlayerNum & " has validated Number " & intComputerPlayerChoice & ", "  & strRowNames( intComputerPlayerChoice-1 ) )                        'new
  end if 	 
end sub

sub CreateRadioButton(strRBName, Value, Label)
    PrHTML ("<input type=" & chr(34) & "radio" & chr(34) & " name=" & chr(34) & strRBName &      chr(34) & " value=" & Chr(34) & Value & Chr(34) & "> " & Label & " <br> ")
    
end sub 

sub updateCardsPlayed
  'Reset CardsPlayed session array based on the values in the radio buttons
   dim strCardsPlayedUpdate                      
   strCardsPlayedUpdate = Session("CardsPlayed")   
   dim intPlayerChoice

   intHumanPlayerChoice = Request.Form("AssmChoice")
   strCardsPlayedUpdate(intHumanPlayerChoice,1) = "X" : ' Added since last valid run June 13th 
   'PrHTML("Player choose " & intHumanPlayerChoice & "<br>")

   Session("CardsPlayed") = strCardsPlayedUpdate  
end sub

sub ComputerPlayer                                                                              
    ' dim strCardsPlayedUpdate                                                                  
    'strCardsPlayedUpdate = Session("CardsPlayed") 
                                                 
 'FIRST ASSUME THAT THIS IS AN ASSUMPTION CHOICE INSTEAD OF AN OH-BY-THE-WAY CARD
  
   'In the beginning, it will be a random choice from among the assumptions unchecked           
    dim intUnchosen                                                                              
    dim i 
    dim intCardCounter   
    dim intChoice   
    dim intFound    'checks to see if the random number has found the correct row                                 'new 
    
      intFound = 0                                                                                               'new                                                                             
      intUnchosen = 4                                                                             
      'determine how many unchosen assumptions there are left
      for i = 1 to 4                                                                            
        if strCardsPlayed(i,intActPlayerNum) = "X" then intUnchosen = intUnchosen-1             
      next   
      if intUnchosen>0 then                                                                                      
          'prHTML("There are " & intUnchosen & " left for player " & (intActPlayerNum-1) )            
          randomize                                                                             
          intChoice = Randbetween_1_and(intUnchosen)                                                 
    
          'update the cards played by moving to the intChoice-th unchosen position
           intCardCounter = 0                                                                        
           for i = 1 to 4   
             if intfound = 0 then                                                                           'new                                                               
              if strCardsPlayed(i,intActPlayerNum) <> "X" then intCardCounter = intCardCounter +1    
              if intCardCounter = intChoice then 
                  strCardsPlayed(i,intActPlayerNum) = "X"             
                  intComputerPlayerChoice = i                                                               'new 
                  Session("ComputerPlayerChoise") = i                                                       
                  intfound=1                                                                             'new 
               end if                                                                                        
             end if                                                                                     'new    
          next                                                                                      
            
           Session("CardsPlayed") = strCardsPlayed   
                                                  
      Else                                                                                            
          PrHTML("Player " & (intActPlayerNum -1) & " has no choices left")                            
      end if                                                                                           
      PrHTML("Player " & intActPlayerNum & " made choice " & intComputerPlayerChoice)                               'new        
end sub                                                                                        
   
%>


      <input type="submit" name="Submit" value="End Turn" />
  </form>
</body>

</html>

