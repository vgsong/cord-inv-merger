# PDF Merger

### Context:

The Company needs to attach invoices based on work group per client's request. The finance deparment obtains a monthly invoice batch where hundred of invoices gets generated.
Once generated, the finance deparment attachs pre-approved clients pdf support into the invoice if the 

Per client's request, the Company needs to attach pre-approved pdf support file into the invoices. Only specific work group needs to have pre-approved support attached. The finance department obtains a hundreds of invoices from a monthly batch. The invoices contain labors hours incurred in the period. (i.e. hours worked per billable fte)

The script will collect the string from the batch generated invoices using regex. It will collect the invoice numnber, project number(Work Order included), and check if the name exists in the invoice.
If so it creates a work order where the attacher then picks up the files in TS dir and attach the pre-approved support based on the "LAST, FIRST" name

This script supports the finance deparmtment in reducing the manual processes.

