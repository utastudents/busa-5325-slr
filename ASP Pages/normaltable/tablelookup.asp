<html>
<body>
 <%QuestionType = Request.form("GO")
   Randomize
   'Input, check and look up zscore 1
   inZscore = request.form("GivenZscore")
   if isnumeric(inzscore) = true then
       zscore=cdbl(inzscore)
   else
       zscore= round(rnd()*3,2)
   end if
   tablePct = round(cumnor(abs(zscore)),4) -.5 

  'Input, check and look up zscore 1
   inZscore2 = request.form("GivenZscore2")
   if isnumeric(inzscore2) = true then
       zscore2=cdbl(inzscore2)
   else
       zscore2= round(rnd()*3,2)
   end if
   tablePct2 = round(cumnor(abs(zscore2)),4) -.5 

    if questiontye = "A" then
        if zscore < 0 then
           strgraphtype ="J"
        else 
           strgraphtype = "E"
    end if   
    
   if questiontye = "B" then
        if zscore < 0 then
           strgraphtype ="H"
        else 
           strgraphtype = "J"
    end if

   if strGraphType="A" then
         if zscore < 0 then zscore=abs(zscore)
         strHeading = "<img src='gAs.gif'>" 
 
         strAnswer = "Pr ( 0 &lt Z &lt " & zscore & " ) = " & tablepct 
      end if   
   
   if strGraphType="B" then
         if zscore2 > 0 then zscore2=abs(zscore2)
         if -1*(zscore) <> zscore2 then zscore2=-1*(zscore)
                 
          strHeading = "<img src='gBsSolve.gif'>" 
 
         strAnswer = "Pr ( " & zscore2 & " &lt Z &lt " & zscore & " ) = " 
         strAnswer = strAnswer & " 2 x Pr ( 0 &lt Z &lt " & zscore & " ) = "  
         strAnswer = strAnswer & "2 x " & tablepct & " = " 
         strAnswer = strAnswer & round((2*TablePct),4)
      end if   

    if strGraphType="C" then
         if zscore2 < 0 then zscore2=abs(zscore2)
         if (zscore) > 0  then zscore=-1*(zscore)
                 
          strHeading = "<img src='gCsSolve.gif'>" 
 
         strAnswer = "Pr ( " & zscore & " &lt Z &lt " & zscore2 & " ) = " 
         strAnswer = strAnswer & " Pr ( " & zscore & " &lt Z &lt 0 )  + "  
         strAnswer = strAnswer & " Pr ( 0 &lt Z &lt " & zscore2 & " ) = "  
         strAnswer = strAnswer & tablepct & " + " & tablepct2 & " = "  
         strAnswer = strAnswer & round((TablePct+tablepct2),4)
      end if   

       if strGraphType="D" then
         if zscore2 < 0 then zscore2=abs(zscore2)
         if (zscore) < 0  then zscore=abs(zscore)
         if zscore > zscore2 then
             temp=zscore
             temppct = tablepct
             zscore=zscore2    
             tablepct=tablepct2       
             zscore2= temp
             tablepct2 = temppct
         end if  
          strHeading = "<img src='gDsSolve.gif'>" 
 
         strAnswer = "Pr ( Z &lt " & zscore & " or Z &gt " & zscore2 & " ) = " 
         strAnswer = strAnswer & " Pr ( 0 &lt Z &lt " & zscore2 & " ) "  
         strAnswer = strAnswer & " - Pr ( 0 &lt Z &lt " & zscore & " ) = "  
         strAnswer = strAnswer &  tablepct2 & " - " & TablePct & " = " 
         strAnswer = strAnswer &  tablepct2 - TablePct 


      end if   

   

    if strGraphType="E" then
         if zscore < 0 then zscore=abs(zscore)
         strHeading = "<img src='gEsSolve.gif'>" 
         strAnswer = "Pr ( Z > " & zscore & " ) = 0.5 - " & tablepct & " = " 
         strAnswer = strAnswer & round((.5-tablepct),4)  
    end if  
    
    if strGraphType="F" then
         if zscore2 < 0 then zscore2=abs(zscore2)
         if -1*(zscore) <> zscore2 then zscore2=-1*(zscore)
                 
          strHeading = "<img src='gFsSolve.gif'>" 
 
         strAnswer = "Pr ( Z &lt " & zscore2 & " or Z &gt " & zscore & " ) = " 
         strAnswer = strAnswer & " 1.0 - 2 x Pr ( 0 &lt Z &lt " & zscore & " ) = "  
         strAnswer = strAnswer & "1.0 - 2 x " & tablepct & " = " & round((1-2*TablePct),4)
      end if   

    if strGraphType="G" then
         if zscore2 < 0 then zscore2=abs(zscore2)
         if (zscore) > 0  then zscore=-1*(zscore)
                 
          strHeading = "<img src='gGsSolve.gif'>" 
 
         strAnswer = "Pr ( Z &lt " & zscore & " or Z &gt " & zscore2 & " ) = " 
         strAnswer = strAnswer & " 1.0 - Pr ( " & zscore & " &lt Z &lt 0 ) "  
         strAnswer = strAnswer & " - Pr ( 0 &lt Z &lt " & zscore2 & " ) = "  
         strAnswer = strAnswer & "1.0 - " & tablepct & " - " & TablePct2 & " = " 
         strAnswer = strAnswer & round((1 - tablepct - TablePct2),4) 


      end if   

    if strGraphType="H" then
         if zscore < 0 then zscore=abs(zscore)
         strHeading = "<img src='gHssolve.gif'>" 
          strAnswer = "Pr ( Z &lt " & zscore & " ) = 0.5 + " & tablepct & " = "
          strAnswer = strAnswer  &  round((tablepct+.5),4)
    end if  

   if strGraphType="I" then
         if zscore2 < 0 then zscore2=abs(zscore2)
         if (zscore) < 0  then zscore=abs(zscore)
         if zscore > zscore2 then
             temp=zscore
             temppct = tablepct
             zscore=zscore2    
             tablepct=tablepct2       
             zscore2= temp
             tablepct2 = temppct
         end if  
          strHeading = "<img src='gIsSolve.gif'>" 
 
         strAnswer = "Pr ( Z &lt " & zscore & " or Z &gt " & zscore2 & " ) = " 
         strAnswer = strAnswer & " 1.0 - Pr ( 0 &lt Z &lt " & zscore2 & " ) "  
         strAnswer = strAnswer & " + Pr ( 0 &lt Z &lt " & zscore & " ) = "  
         strAnswer = strAnswer & "1.0 - " & tablepct2 & " + " & TablePct & " = " 
         strAnswer = strAnswer & round((1 - tablepct2 + TablePct),4) 


      end if   

    
    strPrompt = "Determine the above area "
    strPrompt = strPrompt & "when Z1 =<INPUT NAME='givenZscore' VALUE='" & zscore & "' "
    if (strGraphType = "G") or (strGraphType = "D") or (strGraphtype = "H")  or (strgraphtype = "I") then
           strPrompt = strPrompt & " size=4 <p> and Z2 =<INPUT NAME='givenZscore2' VALUE='" & zscore2 & "' "
    
    end if
    strPrompt = strPrompt & " size=4><p>"
%>
 <% = strheading%>
 <Form action="cumnor.asp" method="post">
     <%response.write strPrompt%>
     <input Name="go" value="<%=strgraphtype%>" Type = "Hidden">          
     <INPUT TYPE=submit VALUE="Calculate">
     <Input type=reset value="Cancel">

      <p> <%=stranswer%> </p> 
 </form>
<p>
<p>


<%  
   Private function cumnor(arg ) 
'**********************************************************************
'
'    SUBROUINE CUMNOR(X,RESULT,CCUM)
'
'
'                             Function
'
'
'    Computes the cumulative  of    the  normal   distribution,   i.e.,
'    the integral from -infinity to x of
'         (1/sqrt(2*pi)) exp(-u*u/2) du
'
'    X --> Upper limit of integration.
'                                       X is DOUBLE PRECISION
'
'    RESULT <-- Cumulative normal distribution.
'                                       RESULT is DOUBLE PRECISION
'
'    CCUM <-- Compliment of Cumulative normal distribution.
'                                       CCUM is DOUBLE PRECISION
'
'
'    Renaming of function ANORM from:
'
'    Cody, W.D. (1993). "ALGORITHM 715: SPECFUN - A Portabel FORTRAN
'    Package of Special Function Routines and Test Drivers"
'    acm Transactions on Mathematical Software. 19, 22-32.
'
'    with slight modifications to return ccum and to deal with
'    machine constants.
'
'**********************************************************************
'
'
'Original Comments:
'------------------------------------------------------------------
'
'This function evaluates the normal distribution function:
'
'                             / x
'                    1       |       -t*t/2
'         P(x) = ----------- |      e       dt
'                sqrt(2 pi)  |
'                            /-oo
'
'  The main computation evaluates near-minimax approximations
'  derived from those in "Rational Chebyshev approximations for
'  the error function" by W. J. Cody, Math. Comp., 1969, 631-637.
'  This transportable program uses rational functions that
'  theoretically approximate the normal distribution function to
'  at least 18 significant decimal digits.  The accuracy achieved
'  depends on the arithmetic system, the compiler, the intrinsic
'  functions, and proper selection of the machine-dependent
'  constants.
'
'*******************************************************************
'*******************************************************************
'
'Explanation of machine-dependent constants.
'
'  MIN   = smallest machine representable number.
'
'  EPS   = argument below which anorm(x) may be represented by
'          0.5  and above which  x*x  will not underflow.
'          A conservative value is the largest machine number X
'          such that   1.0 + X = 1.0   to machine precision.
'*******************************************************************
'*******************************************************************
'
'Error returns
'
' The program returns  ANORM = 0     for  ARG .LE. XLOW.
'
'
'Intrinsic functions required are:
'
'    ABS, AINT, EXP
'
'
' Author: W. J. Cody
'         Mathematics and Computer Science Division
'         Argonne National Laboratory
'         Argonne, IL 60439
'
' Latest modification: March 15, 1992
'
'------------------------------------------------------------------
      Dim I 
      Dim a(5) 
      Dim b(4) 
      Dim c(9) 
      Dim ccum 
      Dim d(8) 
      Dim del 
      Dim eps 
      Dim half 
      Dim p(6) 
      Dim one 
      Dim q(5) 
      
      Dim sixten 
      Dim temp 
      Dim sqrpi 
      Dim thrsh 
      Dim root32 
      Dim X 
      Dim xden 
      Dim xnum 
      Dim y 
      Dim xsq 
      Dim zero 
      Dim min 
      
 '------------------------------------------------------------------
' Mathematical constants
'
' SQRPI = 1 / sqrt(2*pi), ROOT32 = sqrt(32), and
' THRSH is the argument for which anorm = 0.75.
'------------------------------------------------------------------
     one = 1.
     half = 0.5
     zero = 0.
     sixten = 1.6
     sqrpi = 0.398942280401433
     thrsh = 0.66291
     root32 = Sqr(16)
 '------------------------------------------------------------------
' Coefficients for approximation in first interval
'------------------------------------------------------------------
      a(1) = 2.23525203546068
      a(2) = 161.028231068556
      a(3) = 1067.68948546037
      a(4) = 18154.9812533436
      a(5) = 6.56823379182074E-02
      
      b(1) = 47.2025819046882
      b(2) = 976.098551737777
      b(3) = 10260.932208619
      b(4) = 45507.7893350267
'------------------------------------------------------------------
' Coefficients for approximation in second interval
'------------------------------------------------------------------
      c(1) = 0.398941512088135
      c(2) = 8.88314979438838
      c(3) = 93.5066561321779
      c(4) = 597.2702763948
      c(5) = 2494.53758529037
      c(6) = 6848.19045053628
      c(7) = 11602.6514376473
      c(8) = 9842.71483838398
      c(9) = 1.07655767737202E-08
     
     d(1) = 22.2666880443281
     d(2) = 235.387901782625
     d(3) = 1519.37759940755
     d(4) = 6485.55829826676
     d(5) = 18615.5716408851
     d(6) = 34900.952721146
     d(7) = 38912.0032860933
     d(8) = 19685.42967686
'------------------------------------------------------------------
' Coefficients for approximation in third interval
'------------------------------------------------------------------
     p(1) = 8985340579569.9
     p(2) = 0.127401161160247
     p(3) = 2.22352778706498E-02
     p(4) = 1.42161919322789E-03
     p(5) = 2.91128749511688E-05
     p(6) = 2.30734417649402E-02
     
     q(1) = 1.28426009614491
     q(2) = 0.468238212480865
     q(3) = 6.59881378689286E-02
     q(4) = 3.78239633202758E-03
     q(5) = 7.29751555083966E-05
'------------------------------------------------------------------
' Machine dependent constants
'------------------------------------------------------------------
      eps = spmpar(1) * 0.5
      min = spmpar(2)
'------------------------------------------------------------------
      X = arg
      y = Abs(X)
 
     If (y < thrsh) Then
       '------------------------------------------------------------------
       ' Evaluate  anorm  for  |X| <= 0.66291
       '------------------------------------------------------------------
          xsq = zero
          If (y < eps) Then xsq = X * X
          xnum = a(5) * xsq
          xden = xsq
          For I = 1 To 3
              xnum = (xnum + a(I)) * xsq
              xden = (xden + b(I)) * xsq
          next
          cumnor = X * (xnum + a(4)) / (xden + b(4))
          temp = cumnor
          cumnor = half + temp
          ccum = half - temp
      '------------------------------------------------------------------
        ' Evaluate  anorm  for 0.66291 <= |X| <= sqrt(32)
      '------------------------------------------------------------------
      ElseIf (y < root32) Then
          xnum = c(9) * y
          xden = y
          For I = 1 To 7
              xnum = (xnum + c(I)) * y
              xden = (xden + d(I)) * y
          next
          cumnor = (xnum + c(8)) / (xden + d(8))
          xsq = Int(y * sixten) / sixten
          del = (y - xsq) * (y + xsq)
          cumnor = Exp(-xsq * xsq * half) * Exp(-del * half) * cumnor
          ccum = one - cumnor
          If (X > zero) Then
              temp = cumnor
              cumnor = ccum
              ccum = temp
          End If
      '------------------------------------------------------------------
      ' Evaluate  anorm  for |X| > sqrt(32)
      '------------------------------------------------------------------
      Else
          cumnor = zero
          xsq = one / (X * X)
          xnum = p(6) * xsq
          xden = xsq
          For I = 1 To 4
              xnum = (xnum + p(I)) * xsq
              xden = (xden + q(I)) * xsq
          next
          cumnor = xsq * (xnum + p(5)) / (xden + q(5))
          cumnor = (sqrpi - cumnor) / y
          xsq = Int(X * sixten) / sixten
          del = (X - xsq) * (X + xsq)
          cumnor = Exp(-xsq * xsq * half) * Exp(-del * half) * cumnor
          ccum = one - cumnor
          If (X > zero) Then
              temp = cumnor
              cumnor = ccum
              ccum = temp
          End If

      End If
      If cumnor > 1. Then cumnor = 1.
      If (cumnor < min) Then cumnor = 0.
      If (ccum < min) Then ccum = 0.
'------------------------------------------------------------------
' Fix up for negative argument, erf, etc.
'------------------------------------------------------------------
'----------Last card of ANORM ----------
End Function

Private Function spmpar(I ) 
'-----------------------------------------------------------------------
'
'     SPMPAR PROVIDES THE SINGLE PRECISION MACHINE CONSTANTS FOR
'     THE COMPUTER BEING USED. IT IS ASSUMED THAT THE ARGUMENT
'     I IS AN INTEGER HAVING ONE OF THE VALUES 1, 2, OR 3. IF THE
'     SINGLE PRECISION ARITHMETI' BEING USED HAS M BASE B DIGITS AND
'     ITS SMALLEST AND LARGEST EXPONENTS ARE EMIN AND EMAX, THEN
'
'        SPMPAR(1) = B**(1 - M), THE MACHINE PRECISION,
'
'        SPMPAR(2) = B**(EMIN - 1), THE SMALLEST MAGNITUDE,
'
'        SPMPAR(3) = B**EMAX*(1 - B**(-M)), THE LARGEST MAGNITUDE.
'
'-----------------------------------------------------------------------
'     WRITTEN BY
'        ALFRED H. MORRIS, JR.
'        NAVAL SURFACE WARFARE CENTER
'        DAHLGREN VIRGINIA
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'     MODIFIED BY BARRY W. BROWN TO RETURN DOUBLE PRECISION MACHINE
'     CONSTANTS FOR THE COMPUTER BEING USED.  THIS MODIFICATION WAS
'     MADE AS PART OF CONVERTING BRATIO TO DOUBLE PRECISION
'-----------------------------------------------------------------------
 '     ..
'     .. Local Scalars ..
     Dim b 
     Dim binv 
     Dim bm1 
     Dim one 
     Dim w 
     Dim z 
     Dim emax 
     Dim emin 
     Dim ibeta 
     Dim m 
'     ..
'     .. Executable Statements ..
'
    if i <= 1 then
      b = ipmpar(4)
      m = ipmpar(8)
      spmpar = b ^ (1 - m)
 
    elseif i=2 then
      b = ipmpar(4)
      emin = ipmpar(9)
      one = CDbl(1)
      binv = one / b
      w = b ^ (emin + 2)
      spmpar = ((w * binv) * binv) * binv
    Else
      ibeta = ipmpar(4)
      m = ipmpar(8)
      emax = ipmpar(10)
'
      b = ibeta
      bm1 = ibeta - 1
      one = CDbl(1)
      z = b ^ (m - 1)
      w = ((z - one) * b + bm1) / (b * z)
'
      z = b ^ (emax - 2)
      spmpar = ((w * z) * b) * b
 End if

End Function

Private Function ipmpar(intI ) 
'    IPMPAR IS AN ADAPTATION OF THE FUNCTION I1MACH, WRITTEN BY
'     P.A. FOX, A.D. HALL, AND N.L. SCHRYER (BELL LABORATORIES).
'     IPMPAR WAS FORMED BY A.H. MORRIS (NSWC). THE CONSTANTS ARE
'     FROM BELL LABORATORIES, NSWC, AND OTHER SOURCES.
' Machine constants for ibm pc
   Select Case intI
    Case 1
      ipmpar = 2
    Case 2
      ipmpar = 31
    Case 3
      ipmpar = 2147483647
    Case 4
      ipmpar = 2
    Case 5
      ipmpar = 24
    Case 6
      ipmpar = -125
    Case 7
      ipmpar = 128
    Case 8
      ipmpar = 53
    Case 9
      ipmpar = -1021
    Case 10
      ipmpar = 1024
   End Select
End Function
Private Function devlpl(a() , n , X ) 
'**********************************************************************
'
'     DOUBLE PRECISION FUNCTION DEVLPL(A,N,X)
'              Double precision EVALuate a PoLynomial at X
'
'
'                              Function
'
'
'     returns
'          A(1) + A(2)*X + ... + A(N)*X**(N-1)
'
'
'                              Arguments
'
'
'     A --> Array of coefficients of the polynomial.
'                                        A is DOUBLE PRECISION(N)
'
'     N --> Length of A, also degree of polynomial - 1.
'                                        N is INTEGER
'
'     X --> Point at which the polynomial is to be evaluated.
'                                        X is DOUBLE PRECISION
'
'**********************************************************************
'
'      ..
'     .. Local Scalars ..
      Dim term 
      Dim I 
'     ..
'     .. Executable Statements ..
      term = a(n)
      For I = n - 1 To 1 Step -1
          term = a(I) + term * X
      next
      devlpl = term
      

End Function

Private Function stvaln(p) 
'
'**********************************************************************
'
'     DOUBLE PRECISION FUNCTION STVALN(P)
'                    STarting VALue for Neton-Raphon
'                calculation of Normal distribution Inverse
'
'
'                              Function
'
'
'     Returns X  such that CUMNOR(X)  =   P,  i.e., the  integral from -
'     infinity to X of (1/SQRT(2*PI)) EXP(-U*U/2) dU is P
'
'
'                              Arguments
'
'
'     P --> The probability whose normal deviate is sought.
'                    P is DOUBLE PRECISION
'
'
'                              Method
'
'
'     The  rational   function   on  page 95    of Kennedy  and  Gentle,
'     Statistical Computing, Marcel Dekker, NY , 1980.
'
'**********************************************************************
'     ..
'     .. Local Scalars ..
 
      Dim sign 
      Dim y 
      Dim z  
'     ..
'     .. Local Arrays ..
      Dim xden(5) 
      Dim xnum(5) 
'     ..
'     ..
'     .. Data statements ..
      xnum(1) = -0.322232431088
      xnum(2) = -1.
      xnum(3) = -0.342242088547
      xnum(4) = -0.0204231210245
      xnum(5) = -4.53642210148E-05

      
      xden(1) = 0.099348462606
      xden(2) = 0.588581570495
      xden(3) = 0.531103462366
      xden(4) = 0.10353775285
      xden(5) = 0.0038560700634
'     ..
'     .. Executable Statements ..
      If (p >= 0.5) Then
          sign = -1.
          z = p
       Else
          sign = 1.
          z = 1. - p
      End If
      y = Sqr(-2. * Log(z))
      stvaln = y + devlpl(xnum, 5, y) / devlpl(xden, 5, y)
      stvaln = sign * stvaln
 

End Function

Public Function dinvnr(p ) 

'**********************************************************************
'
'     DOUBLE PRECISION FUNCTION DINVNR(P,Q)
'     Double precision NoRmal distribution INVerse
'
'
'                              Function
'
'
'     Returns X  such that CUMNOR(X)  =   P,  i.e., the  integral from -
'     infinity to X of (1/SQRT(2*PI)) EXP(-U*U/2) dU is P
'
'
'                              Arguments
'
'
'     P --> The probability whose normal deviate is sought.
'                    P is DOUBLE PRECISION
'
'     Q --> 1-P
'                    P is DOUBLE PRECISION
'
'
'                              Method
'
'
'     The  rational   function   on  page 95    of Kennedy  and  Gentle,
'     Statistical Computing, Marcel Dekker, NY , 1980 is used as a start
'     value for the Newton method of finding roots.
'
'
'                              Note
'
'
'     If P or Q .lt. machine EPS returns +/- DINVNR(EPS)
'
'**********************************************************************

       Dim q 
   
'     .. Parameters ..

 
      Dim maxit 
      maxit = 100
      Dim eps 
      eps = 0.0000000000001
      Dim r2pi 
      r2pi = 0.398942280401433
      Dim nhalf 
      nhalf = -0.5
'     ..
 '     ..
'     .. Local Scalars ..
 
      Dim strtx 
      Dim xcur 
      Dim cum 
      Dim ccum 
      Dim pp 
      Dim dx 
      
      Dim I 
      Dim qporq 
'     ..
 '     ..
 '     FIND MINIMUM OF P AND Q
'
      q = 1. - p
      qporq = p < q
      If ((qporq)) Then
         pp = p
      Else
         pp = q
      End If
'
'     INITIALIZATION STEP
'
      strtx = stvaln(pp)
      xcur = strtx
'
'     NEWTON INTERATIONS
'
      For I = 1 To maxit
          cum = cumnor(xcur)
          dx = (cum - pp) / (r2pi * Exp(nhalf * xcur * xcur))
          xcur = xcur - dx
          If (Abs(dx / xcur) < eps) Then
             dinvnr = xcur
             If (Not qporq) Then dinvnr = -dinvnr
             Exit Function
          End If
      next

      dinvnr = strtx
'
'     IF WE GET HERE, NEWTON HAS FAILED
'
      If (Not qporq) Then dinvnr = -dinvnr
      
'     IF WE GET HERE, NEWTON HAS SUCCEDED
'
End Function






%>



     

</body>
<a href="p2.asp">Return to Main Menu</a>
</html>
