# An Analysis of Data with VBA

## Ovreview of Project: Editing code to measure performance on stock data

### Purpose: This analysis will breakdown the edits, or refactoring, made towards the stock market dataset with VBA code. It will describe how refactoring can change a pre-existing code can have certain advantages and disadvantages.

#### Results: The following are the code examples that were used to get the results. The first deliverable that had to be accounted for was to set a tickerIndex and set it to 0. 
<img width="1011" alt="Screen Shot 2022-06-07 at 9 23 23 PM" src="https://user-images.githubusercontent.com/102444078/172764422-6a8a9286-ac8f-4819-9b0c-c19d82722082.png">


The second deliverable was to create arrays for ticker, tickerVolumes, tickerStartingPrices, and tickerEnding Prices. The array for tickerVolume had to be set to a 'Long' data type, while the arrays for tickerStartingPrices and tickerEndingPrices had to be set to a 'Single' data type. 

<img width="609" alt="Screen Shot 2022-06-07 at 9 23 12 PM" src="https://user-images.githubusercontent.com/102444078/172764501-8e587dd3-6377-4ace-ae7c-14fba936cb60.png">


The third deliverable was to create a loop that will be used to initialize tickerVolume to 0 and if the the next row's ticker didn't match, to increse it. 

<img width="509" alt="Screen Shot 2022-06-07 at 9 23 30 PM" src="https://user-images.githubusercontent.com/102444078/172764619-484ee5c6-1838-4bdc-9ae3-74c5ff0bfa7b.png">
 

The fourth deliverable was to create another loop that will loop through the stocker index for the three arrays that were created earlier. Inside the loop, there was an if-then statement to check if the current row is the first row with the selected tickerIndex and if it is, to assign the current closing price to the tickerStartingPrices and tickerEndingPrices variable. 

<img width="733" alt="Screen Shot 2022-06-07 at 9 24 27 PM" src="https://user-images.githubusercontent.com/102444078/172764671-ec80fef6-0825-49ab-82ba-60da964ec771.png">


The fifth deliverable had to do with formatting the data so that it can be easier to read - positive returns were colored green and negative returns were colored red. This helped determine which stocks performed well and which ones didn't. 
<img width="443" alt="Screen Shot 2022-06-07 at 9 24 44 PM" src="https://user-images.githubusercontent.com/102444078/172765327-72b39a51-19d4-420e-99cf-851dab7cd35a.png">

If we take a look at the 2017 vs 2018 stock data, we can see that there were major differences between how the stocks performed. For example, in 2017, almost all stocks except for 'TERP', had a positive performance, but in 2018, almost all stocks except for 'ENPH' and 'RUN', performed negatively. 

<img width="388" alt="Screen Shot 2022-06-07 at 9 21 42 PM" src="https://user-images.githubusercontent.com/102444078/172764893-d88b6b64-d41f-48e5-9260-3afa272a1241.png">

<img width="397" alt="Screen Shot 2022-06-07 at 9 22 56 PM" src="https://user-images.githubusercontent.com/102444078/172764927-bddd002c-550f-43b2-a768-e78c0c77edd1.png">


This goes to show that in just a year, the majority of the stocks had drastically changed, meaning they are volatile and highly risky. However, because we can easily visualize which stocks stayed consisted and had positive returns, it becomes easier for us to determine which stock was the most consistent and least risky. Using the updated and refactored script also tells us how fast the data was read and analyzed. With the refactored code, it is faster than the original, allowing us to gain insight as to how long we can expect results.
<img width="544" alt="Screen Shot 2022-06-07 at 9 22 42 PM" src="https://user-images.githubusercontent.com/102444078/172764828-4d2ad4b4-f4c2-4d3d-8e15-3956491f63f0.png">
<img width="559" alt="Screen Shot 2022-06-07 at 9 21 57 PM" src="https://user-images.githubusercontent.com/102444078/172764842-b0859e64-567c-492f-8b23-b73434cb1649.png">


#### Summary: There are several advantages and disadvantages to refactoring code. One advanatge is that by refactoring, we can clean up the script, allowing for future and current users to have an easier time reading and understanding what has been laid out. A much cleaner and organized code can allow the logic to flow withot there being much errrors. One disadvantaged is that it can become more time-consuming to test out the refactored code. One must consistently test out its functionality to determine whether the changes that are being made allow to code to continue its function. If one refactors correctly however, the code should be able to run just as efficiently if not more. One must understand that regardless of the pros and cons, the main concept to consider when refactoring is that it is use to clean up the code that otherwise might be difficult to comprehend and follow. A code that is more soundly structured and easier to read allows for maintenance to be carried out without much difficulty.
