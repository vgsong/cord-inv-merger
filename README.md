# PDF Merger

### Context:

The finance deparment obtains a monthly invoice batch where hundred of invoices gets generated. If an invoice belongs to a specific project, it requires pre-approved timesheets support by the client attached to it. 

The script will collect the string from the batch generated invoices using regex. It will collect the invoice numnber, project number(Work Order included), and check if the name exists in the invoice.
If so it creates a work order where the attacher then picks up the files in TS dir and attach the pre-approved support based on the "LAST, FIRST" name

This script supports the finance deparmtment in reducing the manual processes.

