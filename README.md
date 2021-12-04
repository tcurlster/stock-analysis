# stock-analysis
Using VBA to summarize stock performance over multiple years.

## Project Overview
The purpose of this analysis is to summarize stock performance over multiple years and provide a high level overview in order to make educated financial decisions for the future. 

## Results
From this analysis, I determined the total daily volume and the yearly return of multiple stocks for 2017 and 2018 solely from the creation of a couple of macros in VBA. After the creation of the inital macros, I was able to refactor the code to consolidate the processes captured in multiple macros into one.

### 2017 Review
! Pre-Refactor 2017

- Overall, the stock performance for 2017 was quite impressive. All but one stock had positive returns and 33% had returns over 100%.  Outside of stock analysis, the first iteration of the VBA code used 2 different macros to complete our review.  One to process the data and the second to format.  The processing of data alone took approx. 2 seconds to complete.

!VBA Challenge 2017

- After taking a second look at the logic, I was able to refactor the code to combine the data processing and the formatting into one process and cut the runtime down to 0.27 seconds.  This improvement allows me to run this code on larger sets of data in the future.

### 2018 Review
! Pre-Refactor 2018

- For 2018, the stocks didn't perform nearly as well as the did in 2017.  Only 2 stocks had postive returns on the year and 5 of the 12 stocks reported over 20% losses.

!VBA Challenge 2018

- Similar to 2017, the refactoring of the code saved us approximately 1.7 seconds allowing for more processing in the future.

## Summary
Refactoring code comes with advantages and disadvantages.  As mentioned above, creating the ability to run more efficient code on larger data sets is a huge advantage.  On the flip side, refactoring can also be very tedious and can be confusing. This just means that the coder refactoring the code needs to have a clear plan of attack prior to editing any of their logic. With a clear plan and good documentation, a coder can make their logic run as smooth and efficiently as possible.

