# Analyis of Stock the Returns for 2017 and 2018
## Overview of the Project
The purpose of this project is to use Excel VBA to conduct an analaysis on the return and daily trade volume of the 12 stocks in our dataset. 

## Results
As the graph below indicates, in 2017 the 12 stocks in the dataset displays largely positive returns marked in green. Among all the 12 stocks, SPWR has th highest daily trading volume while DQ shows the largest gain. The situation takes a drastic turn in 2018, with 10 of the 12 stocks displaying a negative return. Only ENPH and RUN shows a positive return in 2018, and ENPH has the highest daily trading volume. RUN has the largest rate of reuturn in 2018 while DQ, which has the largest rate of return in 2017, displays the largest rate of loss.
![All_Stock_2017](https://user-images.githubusercontent.com/74267597/101285285-57b33200-37b2-11eb-8cc5-521cbaa8393d.PNG)
![All_Stock_2018](https://user-images.githubusercontent.com/74267597/101285368-c4c6c780-37b2-11eb-9d6c-ff4135fb2bce.PNG)

## Summary
Refactoring the code makes the original code more efficient, making itself costs less time and less computing power to run. Also, refactoring improves the organization and clarity of the code with fewer steps, making one's life much easier when revisiting the code. However, refactoring could potentially ruin the purpose of the original code and introduce bugs. It can also be a expansive and time consuming process. 

The original code of the VBA script takes more time to run. For instance, for the analysis in 2018 the original code takes approximately 0.6 seconds to run whereas the refactorered code only takes around 0.1 seconds to produce the same result. The disadvantage of the refactored code is that it uses more loops than the original code, adding complexity to the structure.
