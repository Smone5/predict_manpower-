# Predict Fabrication Shop Manpower

A script I created in a scheduling contract position to predict the manpower for welding fabrication shop in Florida. 

## Problem
The intitial problem was the fabrication shop was using a lot of overtime to finish projects on the time. The company was wondering why so much overtime was being used to complete jobs? 

One of main problems for answering this question was there was no clean data to analyze. The company had been keeping timesheets for projects at least five years back in PDFs, but did not engineer the data to analyzed in software like Excel. In order to solve the problem I needed to do the following steps.

## Steps
1. Teach myself VBA. Up to this point most of my programming background was in Python. However, the IT department did not grant me the license to use Python. So I was forced to adapt and learn VBA for Microsoft Excel instead.
2. Create a VBA script to travel the companies server and sequentially download thousands of timesheets pdfs invidually. 
3. Once the timesheets were downloaded, I needed to sequentially go through them to extract the information I needed for my analysis and combine it into one file
4. With data extracted, cleaned and combined I could finally begin my analysis

## Analsyis
After exploring the data the analysis the data showed the fabrication shop had two problems.
1. They were getting bottlenecked in the welding phase of the shop. 
2. Because they were behind on one project, they had to use overtime to catch on the next one. This was a snowball affect on all projects.

## Solutions
Many potential solutions were dicussed like using technology to improve the effecieny of welding. Adding more space and welders to alleviate the bottleneck. One perspective I took was trying to predict how much manpower would be needed in order 
