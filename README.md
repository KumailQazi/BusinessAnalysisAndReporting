# BusinessAnalysisAndReporting
Business Analysis and Reporting
This project is meant to demonstrate several of the skills we are looking for in the Data Team for the Reporting Business Analyst position based on a real problem we have needed to solve in the past. Some requirements are left intentionally vague so that we can see your architecting and problem solving skills, and to allow you some creative freedom. That being said, do not hesitate to ask us questions about the project.

Project Description
Create a formatted spreadsheet report using Pandas and OpenPyXL for the fake insurance company Insurio Inc using the provided CSV files as input. The Finance department of Insurio Inc needs a new report to track the changes in the transactional premium. They are interested in a report that shows the YTD changes in premium by policy with the following columns:

Column	Description	Formatting
Policy Number	The Policy Number is a sequential number assigned to every policy when bound	No specific formatting needed
Named Insured	The insured’s full name. A policy can have multiple named insureds, in those cases we want to show all of them separated by comma	No specific formatting needed
Effective Date	This is the date when a coverage or insurance contract goes into effect or starts	d/m/YYYY (eg. 1/8/2018)
Transaction Type	The type of transaction associated with the revision: New, Endorsement, Canceled, Reinstated	No specific formatting needed
Policy Fee	Fees associated with a policy	Number with no decimals
Change in Premium	The difference in premium among previous revision and the current one	Two decimals, thousands separated by comma
The Policy Fees need to be removed from flat cancellations. These are policy cancellation transactions where the cancel date is less or equal than the Effective Date.
