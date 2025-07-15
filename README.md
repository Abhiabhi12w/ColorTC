we wanna automate excel using macro , i will make this as a button to generate TestPlan sheet

I will add new sheet to add this button and we have a template for TestPlan sheet within same file with some colors and stuff

-------------------

TestPlan sheet : 

- Input for top fields 
- Navigate to Testcase sheet and pull formtype and corresponding docids
- where docs used is not empty color and mark completed and pass and comment as Geico
- Add comment automatically - pull scenario from Testcase and add prefix , pull formtype from testcase and add suffix 


---------------------

In our automation sheet : 

we first take user input values 

PBI
QA
DATE
RuleName
UpdateDesc
GenName

then we update these values corresponding to in TestPlan sheet like this - Key Value (we dont have GenName in TestPlan sheet)

Next we take values of formtype and respective docid from Testcase sheet

in that sheet we have them like this 

docid		randomvalues 	randomvalues	......SO ONN 	formType
GM0001	randomvalues 	randomvalues	...SO ONN		 UB


so we need these values to be updated in TestPlan sheet 

how - 

where ever there is formtype under Scenario (we have another header with scenarios take care) header there corresponding to it we update doc all ids of that formType under header Docs Used

so if we have 4 UB formmtypes then we put all of its docid in same cell under Docs Used corresponding to that form type in TestPlan

now under scenarios header we update them

how - 

In Testcase sheet we apply filter on A column and filter only numbers (1,2,3...) and then we pull corresponding scenarios under Scenario headers

we have predfeined format 

To verify whether + GenName + is firing when + pulled Scenario + for + formType 




