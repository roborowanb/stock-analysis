# Stock Analysis
## Overview
  Steve is helping his parents actively manage their portfolio, and as friends, we are helping Steve. The workbook we prepared for their family throughout the module was very useful. But now Steve is looking to tinker with the information being pulled and we see this as an opportunity to reevaluate the effenciency of our code. The idea is to alter the data set we are collecting into our table, while creating a tickerindex variable that allows us to create it with one comprehensive loop which reduces our run time.
  
## Results
  First off, the run times were significantly faster. I'm sure Steve will very much like the improved responsiveness of his spreadsheet.  Now in terms of stock performance, we can clearly see that 2017 was a better year for returns than 2018. Where 11/12 stocks increased in value in 2017, only 1/12 increased 2018.   
  
  ![2017](https://github.com/roborowanb/stock-analysis/blob/6c56400f6208b668bc900e251bddc671055417b2/2017.png)
  
  ![2018](https://github.com/roborowanb/stock-analysis/blob/941e513f960ed04f5da22b6ae45ba8a5db1b863d/2018.png)
  
## Summary
1. The advantages of refactoring code is that you don't have to start from scratch. This can save time getting to a finished product. It can also be performing a task that you haven't written before, so while refactoring you are learning how to perform that task. The disadvantage is that its possible that code can be inherited with inefficiencies that you are unaware of, and might have a tough time figuring out. Another disadvantage might be that your refactored code could be mix and matched with the context of which it was being used. You should be careful about pulling sections of code without considering how the rest of the sub is impacting its environment, in order to mimic the environment where you are trying to implement it. In some cases, where you are familiar with the syntax, and it is a small chunk of code, you may want to consider just writing the code from scratch. 
2. As someone with very little VBA experience, and no experience writing loops, refactoring the original VBA script was very useful. I still had trouble mimicking its intended environment, but this was still much easier than doing it from scratch. It saved time, effort, and helped me learn the syntax. 
