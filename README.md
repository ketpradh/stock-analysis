# **Stock-analysis for Steve by Ketaki**

## **Overview of Project**

The purpose of this project is to refactor the VB script created for stock analysis for better performance. 
Performance evaluation is done by comparing the runtimes of the original and refactored VB scripts.

This stock anlaysis will help the user (Steve) to understand the stock performance of the green energy companies for the years 2017 and 2018 using two metrics - 
1.   Yearly return of the companies in the respective year. 
2.   Total daily volume of stocks traded for the companies.

This will help the user to make profitable investment decisions for his clients.


## **Results**

### Analysis of stocks:
* Overall, the stocks of almost all companies performed really well in 2017 compared to 2018 with only one company, TERP having a negative yearly return in 2017.
However, in 2018, most companies had negative yearly return, except ENPH And RUN. 
* Stock Comparison can be found here - [GitHub 1](https://github.com/ketpradh/stock-analysis/blob/main/Stock%20comparsion.PNG)
* RUN became extremely profitable in 2018(from 5.5% to 84%, ~93% increase), TERP as well – (-7.2% to -5%) and total volume was higher.
* The daily volumes of stocks traded was higher in 2017 than in 2018 for most companies. 
* For the period over 2017-2018:
* **Top Gainers**: RUN and TERP
* **Top Losers** : DQ and SEDG

### Analysis of the script:
*  The refactored script performs much faster than the original script, cutting the runtime by more than 75%. This is because of the removal of the nested for loops in the refactored script. 
The nested for loops was running **12 X rowEnd times** vs now it runs only **rowEnd** times, where rowEnd are the number of rows in a given worksheet. 
*  Script Runtime Comparison - [GitHub 2](https://github.com/ketpradh/stock-analysis/blob/main/Runtime%20comparsion.png)
*  The original script was activating output worksheet for every ticker within the outer loop. The flipping between the worksheets can cause slow runtimes.

## **Summary**

### **Advantages of refactored code(general):**
Reference [StackOverflow](https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software)
*  Refactoring code improves performance of programs by reducing runtimes.
*  Refactoring helps find bugs and improves code maintainability.
*  It makes software easier to understand by adding comments and removing dead code.

### **Advantages of refactored script:**
* The runtime improves by approx. 75% or more, hence giving faster results and performance. Since the runtime is faster, the script can handle more data with improved efficiency and time.
* The memory usage is less as it avoids storing temporary values during nested loop runs and unnecessary evaluation of conditionals for the given ticker.

### **Disadvantages of refactored code(general):**
Reference - [StackOverflow](https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software)
*  It is a time consuming and resource intensive task.
*  It can be risky if the application is big and refactoring may break some functionality.
*  It needs good understanding of the code functionality by the developer.

### **Disadvantages of refactored script:**
*  The refactored code made use of more memory variables by using additional arrays.
*  While the refactored script runs faster and has better performance, additional time was needed to validate the functionality and the results. 
 
