# Project 1: Analysing Retailers & Salesman Performance  
**Notes:** 
- The scripts in this project were designed for Python 3 
- The dashboard was designed for Power BI version 
- The original content of the dataset was completely modified and adapted due to data privacy/ownership (with the owner's consent)

## Executive Summary
The tobacco industry is constantly facing governmental pressure that limits traditional tobacco due to health concerns. Therefore, players need to invest in reduced-risk products (e.g.: e-cigarettes, vaping products, heated tobacco, etc...) in order give consumers safer alternatives. Within their retailers channel, market players usually try to increase market share of their new products against traditional ones, by giving out money incentives for retailers' successful leads and/or sales.

The client company is currently awarding money incentives to their retailers network, if they manage to achieve sales objectives set by the company for their new reduced-risk product. 
- Lead of Customer | Lending of Product to Customer | Sale to Customer = 10 points
- 10 points = 5€
- Eligibility: Points are only awarded if Lead | Lending | Sales objectives for the period are achieved respectively for each Retailer

Acting as a freelance consultant, a project divided into 4 main parts is provided to answer the challenge raised by the client company.

## Part 1 - Data Acquisition
### Business Problem: 
Currently, there is a "Master" Template that Sales Representatives of each region use to register all data from Retailers (e.g.: Sales objectives, Leads/Lending/Sales results, etc...). The "Master" Template Excel file has 6 sheets:
- 1 Excel Sheet for Point of Sales (i.e.: Retailer) Lead, 1 Excel Sheet for Point of Sales Lending, 1 Excel Sheet for Point of Sales Sales
- 1 Excel Sheet for Salesman (i.e.: Retailer worker) Lead, 1 Excel for Salesman Lending, 1 Excel Sheet for Salesman Sales  

Each Sales Representative then sends out by email their completed Excel file to a designated Sales Analyst that compiles all information received into a Final Excel file. This workflow is very time consuming for the Sales Analyst as well as very prone to error and the client would like a fast and more efficient solution.  
![4_Data_Acquisiton_Workflow_1](https://user-images.githubusercontent.com/82218642/136635657-49eaccc1-ea4b-42b0-ba69-0ab5900ce2f5.png)

### Solution:
Build a Python Script that automates the concatenation of all Excel files into 1 and upload it to the Company's Database, **reducing 2 hours of manual data entry to 1 minute.** 
This leaves more time for Analysis and Decision Making, as well as allowing to setup a Dashboard with the relevant data, as shown in part 3 of this project.  
![4_Data_Acquisiton_Workflow_2](https://user-images.githubusercontent.com/82218642/136635660-208eccaa-1a67-40fc-a2a7-366942714b5c.png)

Note that Data Cleaning was provided to the client as part of the final solution in the Python Script, however, for the purpose of simplification for this Repository, that part was included at the beginning of Part 2 - Data Analysis.

## Part 2 - Data Analysis
### Business Problem:
The client company acknowledges this is an expensive program, but considers it vital for their strategy to help consumers transition from traditional tobacco towards healthier reduced-risk tobacco products. In addition to already having a distribution network setup, the Retailers' channel accounts for the majority of traditional tobacco sales.

The issue is that there is a lack of organizational resources to deep dive into the historical data, however, the client would like to **derive insights from the past months it has been running the program, in particular, the return on investment (ROI) of their money incentives campaigns and if it needs to be fine-tuned.**

### Solution:
To begin with, the data is cleaned and **KPIs are identified and/or created** in order to allow for a holistic analysis: 
- **Cost:** Money incentive given to Retailer X in € 
- **Coverage %:** Number of Sales Achieved per Retailer X / Sales Objective per Retailer
- **Average Sales Increase:** Average Sales Increased per Retailer 
- **Incentive Increase Cost:** Cost per Retailer X / Average Sales Increase per Retailer

Subsequently, an **Exploratory Data Analysis is conducted in order to analyse historical data and give meaningful insights** in order to conclude if the strategy has had a positive or negative impact on the company's sales strategy.

## Part 3 - Power Bi Dashboard (TO BE COMPLETED)
### Business Problem:
The client lacks a tool that would allow all resources of the company to quickly analyse development of results on a daily basis in order to maximize decision making. 

### Solution:
An **Exploratory Data Analysis is provided in the form of a Power BI Dashboard is proposed**, available to all the company, allowing for increased transparency and optimized results follow-up. 
Power BI is chosen, since it is already and can be integrated to their current Visualization Dashboards and Infrastructure.

## Part 4 - Final Presentation
All the findings and results from the parts above are compiled and shown with a storytelling and data visualization skills to highlight relevant and workflows to the client.
The presentation is divided in 4 parts that follow a similar structure to the one presented here.
