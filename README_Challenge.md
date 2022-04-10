# Stock Analysis Challenge
## Refactor VBA code and measure performance
## Project Overview/Purpose 
The purpose of this analysis is to make the code work well for a large amount of stocks, in case I later need to expand the dataset to include the entire stock market over the last few years. I will refactor the code I made previously to loop through all the data one time 
in order to collect the same information that I did previously. Then, I’ll determine whether refactoring my code successfully made the VBA script run faster. 

  The ultimate goal here is to make the code more efficient and easier for future users to read. Efficiency can mean that the code takes fewer steps, uses less memory or has better logic. 
  
## Results: 

### Stock Performance Comparison
One could say that 2017 was a Bull year and 2018 was a Bear year, as most stocks went up in value in 2017, and most stocks went down in value in 2018. 
TERP was the only stock on the list that was down in 2017 (by approximately 7%).
ENPH and RUN were the only stocks on the list to go up instead of down in 2018, and spectacularly so. ENPH went up by almost 82% and 
RUN went up by 84%. Looking deeper into the stocks, ENPH and RUN are both solar energy companies. One could assume that both 2017 and 2018 were good years for solar – perhaps the government offered some kind of subsidy or rebate for using solar energy, and those two companies made significant revenue.

<img width="380" alt="VBA_Challenge_2017_Page_2" src="https://user-images.githubusercontent.com/95447175/162641956-7b3055e4-1626-4295-a81c-4c5b50cbc7f2.png">

<img width="389" alt="VBA_Challenge_2018_Page_2" src="https://user-images.githubusercontent.com/95447175/162641959-ceef579d-08ad-4f50-abfa-3dd270c662da.png">

### Original code run times
<img width="247" alt="Old code run time 2017" src="https://user-images.githubusercontent.com/95447175/162641893-6d85ba03-6224-40ef-b61b-6ea85d978dac.png">

  <img width="254" alt="Old code run time 2018" src="https://user-images.githubusercontent.com/95447175/162641896-9168da40-3c0a-4e76-9267-a8f539481513.png">

### Refactored code run times 
<img width="253" alt="VBA_Challenge_2017_Page_1" src="https://user-images.githubusercontent.com/95447175/162641948-50f3425d-cd82-4896-b82f-e5636fc9ccf4.png">

<img width="258" alt="VBA_Challenge_2018_Page_1" src="https://user-images.githubusercontent.com/95447175/162641950-b5946691-a47d-4292-9ec9-b59fc9753862.png">

From the screenshots above one can see that the refactored code run times are much faster than the original code run times (approx. .11 for refactored code versus .55 for original code). One important note is that when the file is first opened and the code is run for the first time, it takes Excel a bit longer to run the macro. For example, for the original code the run times were closer to .66 when I first ran the macro. 

## Summary
From doing this exercise, I could see the advantages of refactoring code. The most imporant advantage being the faster run time, which I can imagine would be very helpful in a business setting where time is of the essence (time is money). Refactoring also saves time for the developer or data analyst making the code, and enables these professionals to get more work done by working smarter instead of harder. 

   One possible disadvantage would be the potential of losing the original code if it's not copied or committed to a git repository. 
In my opinion it's useful to have the original code saved somewhere in order to compare it to refactored code, and think of other ways to improve it, or other ways of using it for different business scenarios. 

  These pros and cons I just mentioned apply directly to the refactored VBA script in that I made sure to commit every change to the code, in order
to have multiple versions saved in my github repository. This way, I could go back and download a previous version of the VBA script in case I broke the code while trying to do something new with it. Regarding the advantage of faster tun times, I could plainly see this in my refactored VBA script as it ran for .11 time versus .55 time for the original code. 
