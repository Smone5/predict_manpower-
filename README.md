# Predict Fabrication Shop Manpower

A script I created in a scheduling contract position to predict the manpower for welding fabrication shop in Florida. 

## Problem
The initial problem was the fabrication shop I working at was using a lot of overtime to finish projects on the time. The company was wondering why so much overtime was being used to complete jobs? 

One of main problems for answering this question with data was there was no clean data to analyze. The company had been keeping timesheets for projects at least five years back in PDFs, but did not engineer the data to analyzed in software like Excel. In order to solve the problem, I needed to do the following steps.

## Steps
1. Teach myself VBA. Up to this point most of my programming background was in Python. However, the IT department did not grant me the license to use Python. So, I was forced to adapt and learn VBA for Microsoft Excel instead.
2. Create a VBA script to travel the company’s server and sequentially download thousands of timesheets PDFs individually. 
3. Once the timesheets were downloaded, I needed to sequentially go through them to extract the information I needed for my analysis and combine it into one file
4. With data extracted, cleaned and combined I could finally begin my analysis

## Analysis
After exploring the data, the analysis showed the fabrication shop had two problems.
1. They were getting bottlenecked in the welding phase of the shop. 
2. Because they were behind on one project, they had to use overtime to catch on the next one. This was a snowball effect on all projects.

## Solutions
Many potential solutions were discussed like using technology to improve the efficiency of welding. Another possible solution was to add more space and welders to alleviate the bottleneck. One perspective I took was trying to predict how much manpower would be needed in order to meet demand based off historical performance. This ended up being a rather simple machine learning model. Using linear regression, I was able to train a linear regression model on historical data to predict with roughly 90% accuracy how much manpower overall would be needed in the fabrication shop based on the number of concurrent projects. 

## Overall Experience
Unfortunately, I was not able to expand on my linear regression model as a new job opportunity came up. However, this was great experience for me in learning how to extract, clean and combine data for analysis. I got explore the data in a lot of different ways like heat maps and used statistical methods to evaluate the correlation of certain variables. The most important lesson I learned was to focus on the business problem. It is easy use the latest machine learning technique in order to impress counterparts. The most important thing is to stay on track with original business problem and solve it in the most efficient and accurate way.

