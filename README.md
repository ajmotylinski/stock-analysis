# Overview
The code that was originally written completed the necessary jobs but it was something that did not scale well. If the dataset was expanded the inefficient code would get the desired results, but the user may end up waiting for the code to run. By taking the steps to refactor the code, it will run faster and still complete the required tasks.

This project includes an analysis of the efficiency of the code and will show that taking the steps to refactor does indeed decrease the time the user is waiting for the code to run.

# Results
The results of refactoring the code was positive. I was able to confirm that refactored code ran quicker and produced the same results. The original code took around 1 second to run for each of the years that it was run for. While the refactored code ran in 0.15 seconds. This resulted in around 85% time savings in the refactored code.

### Original code run for 2017
![Original code run for 2017](/resources/original%202017.png)
### Refactored code run for 2017
![Original code run for 2017](/resources/VBA_Challenge_2017.png)

### Original code run for 2018
![Original code run for 2017](/resources/original%202018.png)
### Refactored code run for 2018
![Original code run for 2017](/resources/VBA_Challenge_2018.png)

# Summary

## Advantages of refactoring code
There are many advantages of refactoring code. The efficiency of the code is a reason that code is refactored. Another benefit of refactoring code is that it can be made more general so the code could be used in more ways. In some cases, the refactored coded may be shorter and simplier which leads to code that is easier to maintain.

## Disadvantages of refactoring code
Refactoring code may take time away from developing new features. This may be viewed as a disavantage of refactoring code. Sometimes the refactored code may be more complex than the original. Depending on the development team, this may lead to challenges in maintaining the code.

## Application to original code
In this case, taking the time to refactor the code reduced the time it took the code to run for a particular dataset. There was a little more complexity introduced with the refactored code since we were using arrays to hold the start, end, and volumes for each of the stocks. This complexity was a good trade off since we saw significant efficiency gains.
