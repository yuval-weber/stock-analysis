# STOCK ANALYSIS WITH VBA + EXCEL

## OVERVIEW: VBA Stock Analysis Project

### Purpose
This project edits, or refactors, the Stock Market Dataset with VBA scripts to loop through all the data one time to collect an entire dataset. The intent of the exercise is to allow the VBA script to run faster through refactoring the code to ensure greater efficiency: fewer steps, less memory, and improving the logic to make it easier for future users to read. 

### Analysis and Challenges
The assignment centered on a stock analysis scenario and outlined a number of challenges: 

- Prepare the dataset `VBA_Challenge.vbs` file for the project.
- Create a Resources folder in **GitHub** to hold the screenshots for run-time pop-up messages after running refactored analyses for 2017 and 2018.
- Create and convert an `XLSM` file from `*.vbs` dataset.
- Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
- Use the steps **Refactor VBA code and measure performance** to add code where indicated by the numbered comments in the starter code file.
- Utilize knowledge of VBA and the provided starter code to refactor the VBA Script dataset to loop through the data one time and collect all of the requested information for analysis.

## RESULTS: Refactor VBA Code and Measure Performance
 
### Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:

**1. The `tickerIndex` is set equal to zero before looping over the rows.**

> Created a `tickerIndex` variable and set it equal to zero before iterating over all the rows. Will use this `tickerIndex` to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.

![The tickerIndex](https://user-images.githubusercontent.com/78886925/110176417-3eca1500-7dd1-11eb-9d03-72c38bbc5ab0.png)

**2. Create arrays for `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.**

> Created three output arrays: `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.
> In the VBA code, the `tickerVolumes` array should be a **Long** data type while the `tickerStartingPrices` and `tickerEndingPrices` arrays should be a **Single** data type.

![Created Arrays](https://user-images.githubusercontent.com/78886925/110176484-5acdb680-7dd1-11eb-99cb-d226a4434a35.png)

**3. The `tickerIndex` is used to access the stock ticker index for the `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` arrays.**

> Created a for loop to initialize the `tickerVolumes` to **zero**. 
> If the next row’s ticker doesn’t match, increase the `tickerIndex`.
![The tickerIndex](https://user-images.githubusercontent.com/78886925/110176417-3eca1500-7dd1-11eb-9d03-72c38bbc5ab0.png)

![Looping the tickerIndex](https://user-images.githubusercontent.com/78886925/110177829-8ce01800-7dd3-11eb-8987-5a72b1de3246.png)

![The tickerIndex](https://user-images.githubusercontent.com/78886925/110162744-23550f00-7dbd-11eb-9f55-23b6d4a2d5d0.png)


**4. The script loops through stock data, reading and storing all of the following values from each row: `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.**

> Created a **loop** that will loop over all the rows in the spreadsheet.
> Inside the **loop** created a script that increases the current `tickerVolumes` **(stock ticker volume)** variable and adds the ticker volume for the current stock ticker.

![Script loops](https://user-images.githubusercontent.com/78886925/110178488-87370200-7dd4-11eb-9ffd-e1f4f658a6cf.png)

**Stored values from** `tickerStartingPrices` **and** `tickerEndingPrices`

> Created an **if-then** statement to check if the current row is the first row with the selected `tickerIndex`. If it is, then assign the current closing price to the `tickerStartingPrices` and `tickerEndingPrices` variable.

![Script loops 2](https://user-images.githubusercontent.com/78886925/110178652-c6655300-7dd4-11eb-830e-29fa60eb2fdf.png)

**5. Code for formatting the cells in the spreadsheet is working.**

> Make positive returns green and negative returns red to visually determine which stocks did well and which ones did not. Added some formatting based on the values of the returns. 

![Code for formatting the cells in the spreadsheet is working](https://user-images.githubusercontent.com/78886925/110178754-e8f76c00-7dd4-11eb-9949-ba77b416176a.png)


**6. Comments to explain the purpose of the code.**

> **Comments** consistent with **Best Practices for Writing Super Readable Code** (https://www.topcoder.com/coding-best-practices/) 

- Commenting & Documentation, 
- Consistent Indentation, 
- Avoid Obvious Comments. 
- Code Grouping,
- Consistent Naming Scheme,
- DRY (Don't Repeat Yourself) Principle, 
- Avoid Deep Nesting,
- Limit Line Length, etc...

![Comments to explain the purpose of the code](https://user-images.githubusercontent.com/78886925/110178920-2cea7100-7dd5-11eb-9c0b-a5a60fd8e9c0.png)


**7. The outputs for the 2017 and 2018 stock analyses in the `VBA_Challenge.xlsm` workbook match the outputs from the AllStockAnalysis in the module**

The stock analysis confirms that the stock analysis outputs for 2017 and 2018 are the same as the dataset example provided. In addition to the Resources folder, the screenshots below show the final Stock Analysis Results, **Final VBA Analysis 2017 and 2018**, and the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. 

***Dataset Examples Provided***

![Dataset examples provided](https://user-images.githubusercontent.com/78886925/110178999-51dee400-7dd5-11eb-8ab6-0ffb61625412.png)

> Below are the Final VBA Analyses.

***Final VBA Analysis 2017***

![VBA_Challenge_2017](https://user-images.githubusercontent.com/78886925/110179163-97031600-7dd5-11eb-9dc5-0a078c7c323a.png)

***Final VBA Analysis 2018***

![VBA_Challenge_2018](https://user-images.githubusercontent.com/78886925/110179167-98344300-7dd5-11eb-93c5-0502db41c276.png)

**8. The pop-up messages showing the elapsed run time for the script are saved as `VBA_Challenge_2017.png` and `VBA_Challenge_2018.png`**

> Running the 2017 and 2018 data stock analyses produced the below elapsed run times for each year.

***Time on VBA_Challenge_2017.PNG***

![Time for 2017 analysis](https://user-images.githubusercontent.com/78886925/110179255-bc901f80-7dd5-11eb-8807-7cbb61a0ce1f.png)

***Time on VBA_Challenge_2018.PNG***

![Time for 2018 analysis](https://user-images.githubusercontent.com/78886925/110179259-be59e300-7dd5-11eb-8eff-1b5519cc629d.png)


## SUMMARY: Our Statement:

### Deliverable with detail analysis:
**1. What are the advantages or disadvantages of refactoring code?**

Code refactoring, much like all coding in my early experience, is about taking small steps and making small changes so improve the product incrementally. I imagine that if I were a finance professional, I'd develop the more skills and experience to do more sophisticated analysis and be able to look for better opportunities in the stock market. Finally, I would imagine that if I were running scripts on billions or trillions of pieces of data, then the time savings would be worth the effort.

**Disadvantages:**

> - Long procedures may contain repetitive code, so it would likely be better to deploy better logic for more efficient code. 
> - A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
> - Complex unstructured code is usually best split in several functions. 
> - The Refactoring process can affect the testing outcomes. 

**Advantages:**
> - Logical errors easily appear in well structure code that contains nested conditionals and loops. 
> - VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.

**2. How do these pros and cons apply to refactoring the original VBA script?**

> Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. To do so by cleaning and making simpler the code is a task well worth pursuing.
