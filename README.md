# Predict Fabrication Shop Manpower

VBA scripts created in a scheduling contract position to predict the manpower for welding fabrication shop in Florida. The data used for this analysis is not shared to protect the privacy of the company. 

## Problem
The initial problem was the welding fabrication shop I working at was using a lot of overtime to finish projects on the time. The company was wondering why so much overtime was being used to complete jobs.  

One of main problems for answering this question with data was there was no clean data to analyze. The company had been keeping timesheets for projects at least five years back in PDFs, but did not engineer the data to analyzed in software like Excel. In order to solve the problem, I needed to do the following steps.

## Steps
1. Teach myself VBA. Up to this point most of my programming background was in Python. However, the IT department did not grant me the license to use Python. So, I was forced to adapt and learn VBA for Microsoft Excel instead.
2. Create a VBA scripts to travel the company’s server and sequentially download thousands of timesheets PDFs individually.
    * [Get List of Projects on Server](https://github.com/Smone5/predict_manpower-/blob/master/Get%20List%20of%20Projects%20On%20Server.txt)
    * [Download Timesheets](https://github.com/Smone5/predict_manpower-/blob/master/Download%20Timesheets%20VBA%20Script.txt)
    * [Find and Delete Any Blank Timesheets](https://github.com/Smone5/predict_manpower-/blob/master/Find%20Blank%20Timesheets%20in%20Folder.txt)
    
3. Once the timesheets were downloaded, I needed to sequentially go through them to extract the information I needed for my analysis and combine it into one file
    * [Combine Timesheets](https://github.com/Smone5/predict_manpower-/blob/master/OpenFile%20VBA%20Script.txt)
    
4. With data extracted, cleaned and combined I could finally begin my analysis

## Analysis
After exploring the data with a sample size 609 working days from 01/02/2013 to 12/31/2014 , the analysis showed the fabrication shop had two problems.
1. They were getting bottlenecked in the welding phase of the shop. 
2. Because they were behind on one project, they had to use overtime to catch on the next one. This was a snowball effect on all projects.

## Solutions
Many potential solutions were discussed like using technology to improve the efficiency of welding. Another possible solution was to add more space and welders to alleviate the bottleneck. One perspective I took was trying to predict how much manpower would be needed in order to meet demand based off historical performance. This ended up being a rather simple machine learning model. Using linear regression, I was able to train a linear regression model on historical data to predict with roughly 90% accuracy how much manpower overall would be needed in the fabrication shop based on the number of concurrent projects.  

![chart](https://github.com/Smone5/predict_manpower-/blob/master/chart.png)

## Possible Errors
Looking back, I realize my linegar regression my prediction was only done on training data and not sample data. It is likely not as accurate as I predicted. Also most of the data came from projects working at one company with a lot of routine work. During my tenure the company was trying diversify their project portfolio. It is likely in the future that linear regression model will not be as accurate predicting the new type of work. I also did not generate a good hypothesis to test against at the very beginning of analysis. As my analysis went on I changed the hypothesis to match the data. I also assumed that space was not a problem in the fabrication shop. While it may seem possible to scale up to 150 workers to meet demand, in reality the fabrication shop does not have enough space or tools for 150 workers. If the shop were to scale up to 150 workers, it is likely the new workers would not be as productive as the current workers either. The model does not account for that.  

## Overall Experience
Unfortunately, I was not able to expand on my linear regression model as a new job opportunity came up. However, this was great experience for me to learn how to extract, clean and combine data for analysis. I got explore the data in a lot of different ways like heat maps and used statistical methods to evaluate the correlation of certain variables. The most important lesson I learned was to focus on the business problem and frame it in a testable hypothesis.

