Attribute VB_Name = "brylbryl1ChangeNameNotNeeded"
Option Explicit

'Benjamin Sumilhig 3-3-2015
'Amo gealidan ang mga integer nga datatype to long datatype
'para dili na siya mag overflow sa running balance ug total payment

Public Sub auditPayment(LoanID As Long)

        '<EhHeader>
        On Error GoTo auditPayment_Err

        TxtLog "Entered auditPayment"

        '</EhHeader>

        'Check sa gallardo kung open na ba ang loan. haha

100     If rsPayment.State = 1 Then rsPayment.Close
102     rsPayment.Open _
                "Select * from tblPayment where (status = 'Good' or status = 'Full Paid') and LoanID = " _
                & LoanID & " order by Date asc, ID asc "

104     If rsLoan.State = 1 Then rsLoan.Close
106     rsLoan.Open "Select * from tblLoan where LoanID = " & LoanID

108     If rsCustomer.State = 1 Then rsCustomer.Close
110     rsCustomer.Open "Select * from tblCustomer where Code = " & rsLoan!code & ""

112     If rsPayment.RecordCount <> 0 Then
            'listCheck.AddItem ("Payments were made...")
114         rsPayment.MoveFirst

            Dim ctr2           As Long

            Dim runningBalance As Long

            Dim totalPayments  As Long
                    
            'check utro gallardo kung open ba ang rsloan
                    
116         runningBalance = rsLoan!principal + (rsLoan!principal * 0.2) - _
                    rsLoan!collection
118         totalPayments = rsLoan!collection
                
120         For ctr2 = 1 To rsPayment.RecordCount
122             runningBalance = runningBalance - rsPayment!paymentsMade
124             totalPayments = totalPayments + rsPayment!paymentsMade
                    
126             If runningBalance <= 0 Then
128                 rsPayment!Status = "Full Paid"
130                 rsPayment!NewBalance = 0
                Else
132                 rsPayment!Status = "Good"
134                 rsPayment!NewBalance = runningBalance
                End If
                        
136             rsPayment!TotalPayment = totalPayments
138             rsPayment.Update
                        
140             If ctr2 = rsPayment.RecordCount Then
142                 rsLoan!TotalPayment = totalPayments
                            
144                 If runningBalance <= 0 Then
146                     rsLoan!TotalAmortization = 0
                    Else
148                     rsLoan!TotalAmortization = runningBalance
                    End If
                                    
150                 rsLoan.Update
                            
152                 rsCustomer!Balance = rsLoan!TotalAmortization
154                 rsCustomer.Update
                End If
                       
156             rsPayment.MoveNext
                'sa movenext ka. hahah.
158         Next ctr2

        Else

160         rsLoan!TotalAmortization = (rsLoan!principal + ((rsLoan!principal) * 0.2)) _
                    - rsLoan!collection
162         rsLoan.Update
164         rsCustomer!Balance = (rsLoan!principal + ((rsLoan!principal) * 0.2)) - _
                    rsLoan!collection
166         rsCustomer.Update
        End If

        '<EhFooter>

        TxtLog "Exited auditPayment"

        Exit Sub

auditPayment_Err:
        ErrReport Err.Description, _
                "LendingClient.brylbryl1ChangeNameNotNeeded.auditPayment", Erl

        Resume Next

        '</EhFooter>

End Sub
