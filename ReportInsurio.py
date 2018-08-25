import pandas as pd
import openpyxl
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.utils.dataframe import dataframe_to_rows

#Opening the workbook
book = Workbook()
sheet = book.active

#Reading csv files
revision = pd.read_csv('revisions.csv',index_col=0)
policyholders = pd.read_csv('policyholders.csv',index_col=0)
policies = pd.read_csv('policies.csv',index_col=0)
fees = pd.read_csv('fees.csv',index_col=0)

#Creating DataFrames
df = pd.DataFrame(revision)
df1 = pd.DataFrame(policyholders)
df2 = pd.DataFrame(policies)
df3 = pd.DataFrame(fees)

#Creating Master Table
result = pd.merge(df, df1, on='revisionId')
result = pd.merge(result, df3, on='revisionId')
result = pd.merge(result, df2, on='policyId')
result1 = result

#Creating blank dataframe for required report
result2 = pd.DataFrame()
result2['PolicyNumber'] = result1['policyNumber']
result2['policyholderName'] = result1['policyholderName']
result2['TransactionType'] = result1['revisionState']
result2['EffectiveDate'] = result1['createDate']
result2['ChangeInPremium'] = result1['writtenPremium']
result2['PolicyFee'] = result1['writtenFee']

#Sorting by attribute written fee (default ascending order) and removing all duplicate records for revisionid.
result1DupRem = result1.sort_values('writtenFee').drop_duplicates(subset='revisionId', keep='first')
result2DupRem = result1DupRem
result2DupRem['createDate'] = pd.to_datetime(result2DupRem.createDate)
result2DupRem['cancelDate'] = pd.to_datetime(result2DupRem.cancelDate)

#Identifying Flat Cancellation when cancel date is <= effective date and making it 0.
result2DupRem.loc[result2DupRem.cancelDate<=result2DupRem.createDate, 'writtenFee'] = "0"

#Sorting data by Create Date.
result2DupRem = result2DupRem.sort_values('createDate')

#Change in Premium group by 'Policy Number' sorted by 'Date Created' and 'Written Premium' for 2nd 'RevisionId' minus 'Written Premium' for 1st 'RevisionId' and so on.
result2DupRem['derivedPremium'] = (result2DupRem.groupby(['policyNumber'])['writtenPremium'].diff().fillna(result2DupRem['writtenPremium']))

#Setting dataformat for change in premium, create date and sorting by policy number and createdate.
result2DupRem['derivedPremium'] = result2DupRem['derivedPremium'].map('{0:,}'.format)
result2DupRem = result2DupRem.sort_values(by=['policyNumber','createDate'])
result2DupRem['createDate'] = result2DupRem['createDate'].dt.strftime('%d/%m/%Y')

#Final data frame for required report with required labels.
finalResult = pd.DataFrame()
finalResult['Policy Number'] = result2DupRem['policyNumber']
finalResult['Named Insured'] = result2DupRem['policyholderName']
finalResult['Transaction Type'] = result2DupRem['revisionState']
finalResult['Effective Date'] = result2DupRem['createDate']
finalResult['Change In Premium'] = result2DupRem['derivedPremium']
finalResult['Policy Fees'] = result2DupRem['writtenFee']

#Setting report headers and formatting
sheet['A1'] = 'Insurio Inc'
sheet['A2'] = 'June 2018 Premium Report'
sheet['A3'] = ''
sheet['A4'] = ''

#Writing dataframe result in excel.
for r in dataframe_to_rows(finalResult, index=False, header=True):
	sheet.append(r)

#Saving the workbook
book.save("FinalReport.xlsx")
