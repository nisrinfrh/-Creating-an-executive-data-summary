#### Prepare Excel data for efficient analysis in PowerBi.

## Summary

In Tthis task  we were asked to use different ways to organize and calculate  the data .

The goal was to create a summary showing how a business's profits changed month by

month in Quarter 1 compared to the same time last year.

I needed to make the report easy to read by changing how it looks and

highlighting important results. I also had to use formulas to:

**Create specific totals for each month in Quarter 1 for two different business years using the SUMIF function.**

**Figure out the percentage difference for each month in Quarter 1 of 2023 compared to the same months in 2022.**

**Use a logical function to test the order value and display the correct tax amount.**

## PREPARE DATA

I started by downloading and opening the Microsoft Excel workbook "Quarter One Report.xlsx,"

Containing a single worksheet labeled "Summary." This sheet displayed sales information for specific 

Products over two years, including wholesale and retail prices as well as sales quantities.

Afterward, I adjusted and organized the headings in the Excel sheet. Initially, I selected cell

A4 and inputted the heading "TOTAL Q1 SALES." Then, I chose cell A10 and typed "Q1 MONTHLY TOTALS" as

The heading. Once the headings were added to cells A4 and A10, I applied various formatting options 

To ensure they were visually impactful.

![Screenshot (174)](https://github.com/nisrinfrh/nisrinfrh.github.io./assets/157531427/4e0888d7-585a-40fc-a1b5-c991335aced7)

The entries  in column G was originally in all capital letters, but we needed to change it. 

I devised a formula utilizing the PROPER function, which adjusts text case.

This formula formats text into lowercase while capitalizing the first letter of each word.

The formula's syntax is: 

**=PROPER(G2)**. Upon application, the content in cell H2 transformed into *"Mountain Bikes"*

![Screenshot (175)](https://github.com/nisrinfrh/nisrinfrh.github.io./assets/157531427/67c11702-0fc8-419c-9187-8fbbd5a8410b)

PROSSES DATA 

Create new columns for month and year, I crafted a formula in cell K2 utilizing the MONTH function,

And another in cell L2 using the YEAR function. These formulas extracted the respective components from 

The date in cell J2. It was necessary to copy these formulas down to row 246. The MONTH and YEAR functions,

Categorized under Date and Time, are designed to extract specific elements from a date-formatted cell. 

The syntax for the formula in K2 was

**"=MONTH(J2)"** 

The syntax for the formula in L2 was

**"=YEAR(J2)"**

![Screenshot (176)](https://github.com/nisrinfrh/nisrinfrh.github.io./assets/157531427/d85cd067-55f4-4815-8547-83b067c17f7b)

 Created a standard multiplication formula that multiplied the retail price by the order quantity. 

The syntax for the formula in P2 was

**"=N2*O2"**

In cell Q2,I created a formula using an IF function that calculated  tax .

The IF function had to check if the amount in P2 was over 2000. If it was, then the amount in P2

Had to be multiplied by 5%. If it was not, then cell Q2  should display a 0.

The syntax for the formula in Q2 was

**"=IF(P2>2000,P2*5%,0)"** 



![Screenshot (177)](https://github.com/nisrinfrh/nisrinfrh.github.io./assets/157531427/2a163590-1953-4425-b8bb-619706087d25)

In cell B6, I created a SUMIF formula to sum the sales values for 2022. The sales values were in the 

Range R2 to R246. The criteria range was the range L2 to L246 .then I created a similar formula in 

Cell C6 with the same cell ranges but changed the criteria to 2023. 

The syntax for the formula in B6 was:

**"=SUMIF(L2:L246,2022,R2:R246)"**

The syntax for the formula in C6 was:

**"=SUMIF(L2:L246,2023,R2:R246)"**


![Screenshot (180)](https://github.com/nisrinfrh/nisrinfrh.github.io./assets/157531427/65152c7f-2bfa-41ce-82c6-74027091751a)

*FOR THE YEAR 2022*
I created a SUMIFS formula to calculate the total for January in cell B12, February in cell B13, and March in cell B14. 
 
I added dollar signs to the R and K cell references so that the formula could be copied down.
 
The syntax for the formula in B12 was: 

**"=SUMIFS($R2:$R246,$L2:$L246,"2022",K2:K246,1)"**

The syntax for the formula in B13 was:

**"=SUMIFS($R2:$R246,$L2:$L246,"2022",K2:K246,2)"**

The syntax for the formula in B14 was:

**"=SUMIFS($R2:$R246,$L2:$L246,"2022",K2:K246,3)"**

![Screenshot (181)](https://github.com/nisrinfrh/nisrinfrh.github.io./assets/157531427/b87f0fd0-a445-4641-b479-26d41fb473e6)

*FOR THE YEAR 2023*

I did the same steps above  for the year 2023 in the cells C12, C13 ,AND C14.

The syntax for the formula in C12 was: 

**"=SUMIFS($R2:$R246,$L2:$L246,"2023",K2:K246,1)"**

The syntax for the formula in C13 was:

**"=SUMIFS($R2:$R246,$L2:$L246,"2023",K2:K246,2)"**

The syntax for the formula in C14 was:

**"=SUMIFS($R2:$R246,$L2:$L246,"2023",K2:K246,3)"**

![Screenshot (181)](https://github.com/nisrinfrh/nisrinfrh.github.io./assets/157531427/b87f0fd0-a445-4641-b479-26d41fb473e6)

I created a Percentage difference formula in D6 which showed the percentage by which sales increased in 2023.

To determine the percentage difference between the results for 2022 and 2023, the total for 2022 first had to 

be subtracted from the 2023 total. The result had then to be divided by the result for 2022. 

 The syntax for the formula in D6 was:
 
**"=(C6-B6)/B6"**

I created a similar formula in D12 and copied the calculation in D12 down to D14. 

**"=(C12-B12)/B12"**
**"=(C13-B13)/B1"**
**"=(C14-B14)/B14"**

![Screenshot (183)](https://github.com/nisrinfrh/nisrinfrh.github.io./assets/157531427/0a56af12-2b35-496b-a201-c58eb86c4f55)

### Conclusion 

In this scenario, my assignment involved employing a diverse set of formatting techniques and crafting various

Formulas to generate data columns within a spreadsheet. My objective was to compute tailored totals and 

Transform standard sales data into a concise summary, primed to guide and influence business decisions.




















