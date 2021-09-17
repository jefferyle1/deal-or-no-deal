Attribute VB_Name = "Module1"
'Procedure that opens text file, puts information into label and randomizes the value for each case by swapping values
Public Sub MoneyReader(Arr() As Single, B As Single)
    
    Randomize
    
    Dim X As Integer
    Dim Y As Integer
    Dim F As Integer
    Dim RandomValue As Integer
    Dim Temp As Single
    
    'Reads dollar amounts from text file and stores them in array
    Open App.Path & "\MONEY.txt" For Input As #1
    Do While Not EOF(1)
        Input #1, Arr(X)
        X = X + 1
    Loop
    Close #1
    
    'Each value is formatted into currency
    For Y = 0 To 25
       frmMain.lblMoney(Y).Caption = Format$(Arr(Y), "Currency")
       B = B + Arr(Y)
    Next Y
    
    'The money array values are swapped randomly so that the cases each have a random value from the money array
    For F = 0 To 25
        RandomValue = Int(Rnd * 26)
        Temp = Arr(F)
        Arr(F) = Arr(RandomValue)
        Arr(RandomValue) = Temp
    Next F
        
End Sub

'Procedure that updates labels with information about # of cases in the main interface of the game
Public Sub Status(ByVal Cases As Integer, ByVal RN As Integer)

    Dim CasesToOpen As Integer
    
    'Code to update labels
    frmMain.lblOpened.Caption = "Opened: " & 26 - Cases
    frmMain.lblRemaining.Caption = "Remaining: " & Cases
    
    'Case select to determine number of cases left in each round
    Select Case RN
        Case 1
            CasesToOpen = Cases - 20
        Case 2
            CasesToOpen = Cases - 15
        Case 3
            CasesToOpen = Cases - 11
        Case 4
            CasesToOpen = Cases - 8
        Case 5
            CasesToOpen = Cases - 6
        Case 6 To 10
            CasesToOpen = 1
        End Select
        
    'Code that ensures proper grammar for case / cases
    If RN = 10 Then
        frmMain.lblBriefCase.Caption = "Open your reserved case"
    ElseIf CasesToOpen = 1 Then
        frmMain.lblBriefCase.Caption = "Eliminate " & CasesToOpen & " Case"
    Else
        frmMain.lblBriefCase.Caption = "Eliminate " & CasesToOpen & " Cases"
    End If
    
End Sub

'Procedure that progresses to next round based on # of cases
Public Sub RoundUpdater(ByVal Cases As Integer, RN As Integer, ByVal B As Single, Offer As Single)
    
    'When a certain number of cases remain, the game progresses to the next round
    Select Case Cases
        Case 20
            RN = 2
        Case 15
            RN = 3
        Case 11
            RN = 4
        Case 8
            RN = 5
        Case 6
            RN = 6
        Case 5
            RN = 7
        Case 4
            RN = 8
        Case 3
            RN = 9
        Case 2
            RN = 10
        Case 1
            RN = 11
        End Select
        
    'At the end of each round, the "deal or no deal" menu pops up
    Select Case Cases
        Case 20, 15, 11, 8, 6, 5, 4, 3, 2
            MenuPopup Cases, RN, B, Offer
        End Select
         
    'Updates the round label
    If RN <> 11 Then
        frmMain.lblRound.Caption = "Round " & RN
    End If
    
End Sub

Public Sub MenuPopup(ByVal CaseRemaining As Integer, ByVal RN As Integer, ByVal B As Single, Offer As Single)
    
    Dim NextRoundCases As Integer
    
    'Code that determines the # of cases that have to be opened in the next round
    Select Case RN
        Case 1
            NextRoundCases = 6
        Case 2
            NextRoundCases = 5
        Case 3
            NextRoundCases = 4
        Case 4
            NextRoundCases = 3
        Case 5
            NextRoundCases = 2
        Case 6, 7, 8, 9
            NextRoundCases = 1
        End Select
    
    'Code that prevents the accidental selection of cases during the menus
    frmMain.fraCases.Enabled = False
    
    'Calculating and displaying the offer
    Offer = ((B / CaseRemaining) * (RN - 1)) / 10
    frmMain.lblOffer.Caption = "The bank offer is " & Format$(Offer, "Currency") & "."
    
    'Code to display in the "deal no deal" menu the number of cases left
    frmMain.lblMenuCases.Caption = "You have opened " & 26 - CaseRemaining & " cases."
    
    'Ensures proper grammar for the # of cases left for the "deal no deal" menu
    If RN = 10 Then
        frmMain.lblMenuCasesRemaining.Caption = "There is " & CaseRemaining - 1 & " case left on the board."
    Else
        frmMain.lblMenuCasesRemaining.Caption = "There are " & CaseRemaining - 1 & " cases left on the board."
    End If
    
    'Code that displays the # of cases that have to be opened in the next round
    If RN < 6 Then
        frmMain.lblNumberToOpen.Caption = "You will have to open " & NextRoundCases & " cases in the next round."
    ElseIf RN < 10 Then
        frmMain.lblNumberToOpen.Caption = "You will have to open " & NextRoundCases & " case in the next round."
    ElseIf RN = 10 Then
        frmMain.lblNumberToOpen.Caption = "You will have to open your reserved case next round!"
        frmMain.lblLastOffer.Caption = "Your last offer was " & Format$(Offer, "Currency") & "."
    End If
    
    'Displays the "deal no deal" menu
    frmMain.fraGame.Visible = True
    
End Sub
