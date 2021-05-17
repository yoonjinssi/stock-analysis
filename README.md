# Stock Performance Analysis Refactored
## Overview of Project
####  Using tickerIndex to loop over the list once, we are refactoring the code to reduce the run time.

## Results

######   As we can see from the image below, using tickerIndex to loop through the data decreased the time it takes to analyze the stock data. In order to reduce the run time, we created tickerIndex to loop through the arrays of tickerVolumes, tickerStartingPrices and tickerEndingPrices.
#### *These are times it took when we used original AllStocksAnalysis code*
<img width="253" alt="2017_Original" src="https://user-images.githubusercontent.com/81896860/118428938-679a4080-b685-11eb-8ba5-5d195adcf475.png">
<img width="257" alt="2018_Original" src="https://user-images.githubusercontent.com/81896860/118428958-741e9900-b685-11eb-9495-982a5b2e443d.png">

#### *These are the times it took when we refactored the code using tickerIndex*
<img width="259" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/81896860/118428718-ee025280-b684-11eb-82a8-c1230407f1f7.png">
<img width="260" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/81896860/118428726-f0fd4300-b684-11eb-8c13-a392b059e385.png">

######  In order to create the tickerIndex that will loop through three different arrays, we used this code:
```
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
```

## Summary 
#### By refactoring code, in general, we were able to bring down the time it took. This will help when the data volume gets higher. However, creating tickerIndex and three different arrays takes a long time to think about and there's more chance that we need lots of debugging because of spellings and mistkakes.

#### using original VBA script was simpler to understand for me and it didn't seem like a long time to run the program. Hoever, as the data gets incredibly large, it will be hard to manage the time it takes. 
