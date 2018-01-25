VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'11/10/2017
'issue:
'Zapata "no data found" message issue if collector date hired date is newer
'than the payment date example date hired is nov 7 then payment date in Oct 18.

'Solution:
'was to change the date hired of the collector.
'////////////////////////////////////////////////////////////////////

'12/17/17 7:42pm
'Main code for loan computation :
'- Project1.frm_Loan.SumCharges
'- Project1.frm_Loan.txtPrincipal_Change

'1/1/18 4:55pm
'deleted some files

