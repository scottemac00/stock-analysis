# VBA of Wall Street

## Overview of Project
The purpose of this analysis was to review the code created for Steve and **refactor** it to loop through all the data in the workbook at one time. The goal was code optimization, and I believe we've achieved that. By reviewing how we coded the 'green_stocks' workbook we were able to re-use some code and improve some code, thus improving overall functionality.
## Results
The phrase, "past performance is not an indicator of future returns" is a foundational idea when analyzing stocks, but with a little bit of Visual Basic for Applications coding the information Steve sought for his parent's investments became easier to visualize. Even though this was a project for a fellow financial analyst, I had user experience in mind when preparing the final product for Steve. The optimized experience provides a faster output than previous efforts, and is more visually stimulating.
### Drab to Almost Fab
What began as a relatively ordinary MS Excel workbook that contained a not small amount of stock performance information became a macro-enabled, easy-to-use interface that does a bit more than simply output some stock information for a given year. 

| ![](https://github.com/scottemac00/stock-analysis/blob/0224d5664e57cd7f7ee4535c6b677a8ee876903f/Resources/Original%20DQ%20Analysis%20Sheet.png) | ![](https://github.com/scottemac00/stock-analysis/blob/0224d5664e57cd7f7ee4535c6b677a8ee876903f/Resources/VBA_Challenge_Output_Example.png)                 |
| --------------------------- | ----------------------------- |
| **DQ Analysis Screenshot**  | **VBA Challenge Screenshot**  |

The image on the left depicts the basic output of the original analysis for Steve with one stock. The image on the right depicts the final product (VBA Challenge) output that displays all green stocks. Note the more visually stimulating VBA Challenge output that allows the user to easily determine loss or gain with color coding. 
### What Happened in 20xx?!
When viewed side-by-side, stock performance in 2017 vastly outpaced that of 2018, with the only two solid performers being *ENPH* and *RUN*. 

| ![](https://github.com/scottemac00/stock-analysis/blob/474ab15608f0fcc1c3fa52c9abaf23a6391d1f4a/Resources/VBA_Challenge_2017%20Output%20Screen.png) | ![](https://github.com/scottemac00/stock-analysis/blob/474ab15608f0fcc1c3fa52c9abaf23a6391d1f4a/Resources/VBA_Challenge_2018%20Output%20Screen.png) |
| --------------------------- | ----------------------------- |
| **2017 Performace**  | **2018 Performance**  |

Trading volumes remained high on both of these stocks and Steve might consider recommending these over the previous *DQ* offering.
### Code Performance 
Optimizing the code enabled the output scripts to run approximately 4/10ths to 5/10ths of a second faster, so in theory adding more data will also take less time to compute given the refactored VBA_Challenge file. The 15,000 or so data points per sheet resulted in script runs as follows: 

| ![](https://github.com/scottemac00/stock-analysis/blob/0224d5664e57cd7f7ee4535c6b677a8ee876903f/Resources/Non-Refactored%20Code%20Time%202017.png) | ![](https://github.com/scottemac00/stock-analysis/blob/0224d5664e57cd7f7ee4535c6b677a8ee876903f/Resources/VBA_Challenge_2017.png) |
| --------------------------- | ----------------------------- |
| **Orginal Code Time, 2017**  | **Refactored Code Time, 2017**  |

| ![](https://github.com/scottemac00/stock-analysis/blob/0224d5664e57cd7f7ee4535c6b677a8ee876903f/Resources/Non-Refactored%20Code%20Time%202018.png) | ![](https://github.com/scottemac00/stock-analysis/blob/0224d5664e57cd7f7ee4535c6b677a8ee876903f/Resources/VBA_Challenge_2018.png) |
| --------------------------- | ----------------------------- |
| **Orginal Code Time, 2018**  | **Refactored Code Time, 2018**  |

Additionaly, while you might think adding an additional variable and For loops to our code would mean additional run time, the opposite proved true. Here is an example of the original code loops, followed by the refactored code loops:

| ![]() | ![]() |
| --------------------------- | ----------------------------- |
| **Orginal Code Loops**  | **Refactored Code Loops**  |


## Summary
Advantages to refactoring code include: improved output times, refined user interface and visual outputs, and simplifiying the code for future re-use. Disadvantages might include: time spent on said refactor versus sustaining current products (project resource management), and the potential to break the code or maybe even uneccesarily complicate code at the expense of a more detailed output.
Refactoring the orginal VBA script resulted in a faster output time while enabling a more refined user interface and visually stimulating output. The main disadvantage was the time it took this novice coder to get the improved code to work, so there were several 'defects' before the final 'save'. Overall, with practice, the valued input of classmates, and a bit of grit, Steve (or whomever!) now has a reasonably functional macro-enabled workbook to use in stock analysis.



