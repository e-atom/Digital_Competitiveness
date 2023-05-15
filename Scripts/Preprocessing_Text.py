##########################################################################################################
## Script Name: Preprocessing_Text.py
## Description: This script processes pdf files, extracts the required tables and outputs it to excel
## Author: yuri_bianco
## Date Created: 2023/04/1
## Date Modified: 2023/5/14
## Version: 1.2
##########################################################################################################

#Import Libraries
import pandas as pd
import numpy as np
import pypdf as pdf
import re

##############################Create Functions############################################################

#Create function to read pages:
def read_file(file, pagenumber ,start_line,end_line):
    with open(file, "rb") as f:
        read_file = pdf.PdfReader(f)
        page = read_file.pages[pagenumber]
        text= page.extract_text()
        start_index = text.index(start_line)
        end_index = text.index(end_line)
        table = text[start_index:end_index]
    return table

#Create function to process files
def Preprocessor(table):

    df = pd.DataFrame()

    countries = r'(^\b\w[^0-9]+\w\b)'
    number = r'(-.+|\d.+)'
    

    # Compile the regular expression pattern into a regular expression object
    regex = re.compile(countries, re.MULTILINE)
    regex2 = re.compile(number, re.MULTILINE)


    #Use the findall method of the regular expression object to extract the matches from the table data
    matches = regex.findall(table)
    numbers = regex2.findall(table)

    df['Country'] = matches
    df['Numbers'] = numbers

    # Fix one observation:
    # Split the "Numbers" column on whitespace
    split_numbers = df['Numbers'].str.split()

    # Create a new DataFrame with the split values
    new_df = pd.DataFrame(split_numbers.tolist(), columns=[2018,2019,2020,2021,2022,20181,20191,20201,20211,20221])

    #Convert all scores to int
    cols =[2018,2019,2020,2021,2022,20181,20191,20201,20211,20221]
    new_df[cols]= new_df[cols].apply(pd.to_numeric, errors='coerce', axis=1)

    # Concatenate the new DataFrame with the "Country" column
    result_df = pd.concat([df['Country'], new_df], axis=1)
    Overall_Rank= result_df.iloc[:,0:6]
    Knowledge_Rank = result_df.iloc[:,[0,6,7,8,9,10]]
    Knowledge_Rank.columns = ['Country', 2018, 2019, 2020, 2021, 2022]

    return Overall_Rank, Knowledge_Rank

#Special Function for page 40
def Preprocessor40(table):
    df = pd.DataFrame()

    countries = r'(\b\w[^0-9]+\w\b)'
    number = r'^(.*?)(?=\s+(?:(?:[A-Z][a-z]*\.?\s*)+)$)'

    # Compile the regular expression pattern into a regular expression object
    regex = re.compile(countries, re.MULTILINE)
    regex2 = re.compile(number, re.MULTILINE)


    #Use the findall method of the regular expression object to extract the matches from the table data
    matches = regex.findall(table)
    numbers = regex2.findall(table)

    df['Country'] = matches
    df['Numbers'] = numbers

    # Fix one observation:
    # Split the "Numbers" column on whitespace
    split_numbers = df['Numbers'].str.split()

    # Create a new DataFrame with the split values
    new_df = pd.DataFrame(split_numbers.tolist(), columns=[2018,2019,2020,2021,2022,20181,20191,20201,20211,20221,'None'])


    #Convert all scores to int
    cols =[2018,2019,2020,2021,2022,20181,20191,20201,20211,20221]
    new_df[cols]= new_df[cols].apply(pd.to_numeric, errors='coerce', axis=1)

    #Concatenate and split dataframes
    results_df = pd.concat([df['Country'], new_df], axis=1)
    Tech=results_df.iloc[:,0:6]
    Future_Readiness= results_df.iloc[:,[0,6,7,8,9,10]]
    Future_Readiness.columns = ['Country', 2018, 2019, 2020, 2021, 2022]

    return Tech, Future_Readiness

def Preprocessor_sub(table): 
    df = pd.DataFrame()

    countries = r'(^\b\w[^0-9]+\w\b)'
    number = r'(-.+|\d.+)'


    # Compile the regular expression pattern into a regular expression object
    regex = re.compile(countries, re.MULTILINE)
    regex2 = re.compile(number, re.MULTILINE)


    #Use the findall method of the regular expression object to extract the matches from the table data
    matches = regex.findall(table)
    numbers = regex2.findall(table)

    df['Country'] = matches
    df['Numbers'] = numbers

    # Fix one observation:
    # Split the "Numbers" column on whitespace
    split_numbers = df['Numbers'].str.split()

    # Create a new DataFrame with the split values
    new_df = pd.DataFrame(split_numbers.tolist(), columns=["Talent","Training & education","Scientific concentration","Regulatory framework","Capital","Technological framework"])

    #Convert Numeric features
    cols=["Talent","Training & education","Scientific concentration","Regulatory framework","Capital","Technological framework"]
    new_df[cols]=new_df[cols].apply(pd.to_numeric,errors='coerce',axis=1)
    
    sub_factor = pd.concat([df['Country'], new_df], axis=1)

    return sub_factor

#Create a function to write all dataframes in order to Excel
def ExcelWriter(df1,df2,df3,df4,df5):
    
    # create a writer object for the Excel workbook
    writer = pd.ExcelWriter('IMD_digital_competencies.xlsx')
    
    #Add the files into the Excelsheet
    df1.to_excel(writer, sheet_name='Overall', index=False)
    df2.to_excel(writer, sheet_name='Knowledge', index=False)
    df3.to_excel(writer, sheet_name='Technology', index=False)
    df4.to_excel(writer, sheet_name='FutureReadiness', index=False)
    df5.to_excel(writer, sheet_name='SubFactors', index=False)
    
    # save the workbook
    writer.save()

##########################################################################################

#Execute Functions and Preprocess Table

#Pass read_file function to fetch table
Overall, Knowledge = Preprocessor(read_file(file, 39, "Argentina", "OVERALL"))
Tech, Future_Readiness=Preprocessor40(read_file(file,40,"OVERVIEW2018",'FUTURE READINESS'))
Subfactors=Preprocessor_sub(read_file(file,41,"Argentina", "TECHNOLOGY KNOWLEDGE"))

#Write to Excel
ExcelWriter(Overall,Knowledge,Tech,Future_Readiness,Subfactors)

##########################################################################################