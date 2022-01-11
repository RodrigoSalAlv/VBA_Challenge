# VBA_Challenge

## Overview of Project
### Purpose
The purpose of this project it's to, under the easiest and faster way, found the "ticker" with the higher return by the end of two years: 2017 and 2018.

##Results
###Analysis Results
As we can see in the next two images, 2017 was a very good year for most tickers with an exception for the "TERP" ticker who gets a -7.21% return. In the other hand we can see that for 2017, was a great year for "DQ" ticker, being the one with the higher return over all (Perhaps that's why Steve's parents want to know the DQ's return for 2018).But, by the end of 2018, the ticker gets a negative return (-62.6%).


![image](https://user-images.githubusercontent.com/96214489/148877844-6d717769-26a2-445e-aa69-47a001e9060b.png)
![image](https://user-images.githubusercontent.com/96214489/148877906-8acc95dd-2a28-46a3-a7b2-e16d99f63661.png)

In 2018 the return was very bad, over 80% of tickers gets a negative return by the end of the year. Only the "ENPH" and "RUN" tickers gets a positive return. "ENHP" decrease his procentage in 47.6% and still gets an 81.92% return. For "RUN" was a great year, getting an increase of 78.4% by the end of 2018.

###Script Results
The first script built gets the results we want, but it's slighty slow, the reason of this is because: what the script does is "look for the values in the array, calculate the conditions and print the results for each ticker" this script does not store the values calculated, that's why all of them needs to be printed after the calculations.

![image](https://user-images.githubusercontent.com/96214489/148880908-4b92fdfb-d100-4511-9fdd-05b992352a0f.png)

This makes the macro to take a little bit more to be completed, this are the times that takes for each year:

![image](https://user-images.githubusercontent.com/96214489/148881139-4c0fba1f-f1ff-4909-8431-75e431e0503d.png)

It's totally different with the refactored macro. This one runs faster and, in certain point, the code looks more 'neat'. The use of arrays helps by storing the calculations without printing them, and this are only printed when it's needs to.

![image](https://user-images.githubusercontent.com/96214489/148881793-a982479d-950e-45a3-afcc-fa423512e7a2.png)

The times are much better after refactoring the macro:

![image](https://user-images.githubusercontent.com/96214489/148882093-f081a0b8-1c57-422e-9be1-6a841224d5ad.png)

The disadvantages of this macro are two: 
1. It's necessary to write the name of all the tickers in the "Tickers" array and,
2. The quantity of tickers also needs to be write it down, by the moment of declare the variables.

The macro does not get this information by itself, but, just by adding a few more arrays this can be fixed:

![image](https://user-images.githubusercontent.com/96214489/148882711-cb334ea6-fc2c-4609-aacf-19d1c193b5d4.png)

By doing this little changes, makes the macro more flexible in case tha data change, for example: the quantity of tickers, the name of tickers, the year, etc. We don't need to update the code every time and the time to complete increase just slighty.

![image](https://user-images.githubusercontent.com/96214489/148883886-7b6144fb-ff11-4e65-a115-20390a19ef6d.png)

##Summary

- What are the advantages or disadvantages of refactoring code?
In my point of view, there is no disadvantage when refactoring a code, if this is going to be refactored it's because will be improved. 
Speaking of advantages, this can be many more: the time of excecuting the macro could be less, the code could be easier to read, the lines of code could be less. It's always good to improve something.

-How do these pros and cons apply to refactoring the original VBA script?
In this case there is no cons by the moment of refactoring, the executing time drastically improved and the results were the same. But, as I mention, the macro is unflexible by accepting more data. But wiht another "refactoring" this get fixed.



