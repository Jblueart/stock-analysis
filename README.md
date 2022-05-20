# VBA Challenge: Module 2

### Project Overview

This project focuses on Stock performance in an effort to determine if specific stocks are a good buy or not. Obviously, no one can predict market performance with 100% Accuracy, but we can at least invest with purpose after doing sound research. Through using our analysis tool created in Visual Basic for Applications, otherwise known as “VBA” we can measure specific stock performance for various years with just a few clicks. While learning how to use VBA is complex, it is a powerful tool. My goal is to analyze the stocks utilizing new tools to me, such as Conditional statements, conditional formatting, nested loops & Code-refactoring in an effort to speed up the analysis time & to simplify the investigation. 

#### Biggest Takeaway

One of the main things I realized is that nested loops save tons of time. At first, I thought that I should implement the conditional formatting after running the primary analysis, which was quite fast. The primary analysis took less than a second, however the formatting portion caused a huge strain on my system. It took nearly 6 seconds just to format the results. Granted 7 seconds for an analysis of stocks isn’t bad, it complicated the process unnecessarily. 
 ![image](https://user-images.githubusercontent.com/104408782/169194829-68be8cd5-ee41-4999-8953-727bb8c3698f.png)
![image](https://user-images.githubusercontent.com/104408782/169194843-980e1edd-0ea1-40b3-8cfe-a639dcc39b96.png)

 

After I tried using a nested loop to run the analysis, I realized that it cut down on the run time considerably. My reformatted code combined the formatting with the analysis and saved over 5 seconds. It was so quick to add the formatting as a nested loop as well as simplified the process for the client. One click vs two, and much time saved. The computer didn’t have to struggle as much. It goes to show that the old adage of keep it simple is wise as well as efficient. 
 
![image](https://user-images.githubusercontent.com/104408782/169194868-24078440-a94d-4fe4-aba1-4f13225f1272.png)


### Simplicity

The beauty of using conditional formatting helps give stock performance at a glance. It made it super simple to see the return on investment for the stocks in question. Here you can see that several stocks performed well in 2017, but not quite as well in 2018. This could be due to a number of factors, but it also helps us realize that sometimes the initial performance of a stock is not always indicative of long term performance. As we’ve seen in recent months, the stock market is a fickle beast. 

#### 2017 analysis
![image](https://user-images.githubusercontent.com/104408782/169194885-6baf8323-0291-4b06-a20d-870dd79b3bea.png)

 

#### 2018 analysis
![image](https://user-images.githubusercontent.com/104408782/169194905-bce01dfd-c118-47b1-99b5-854e34d64d3d.png)

 
 ## Update: 
 Ran into a bug with my first code when I tried to add in the starting & ending price because I misunderstood the ruberic. The latest version of my file runs in 1.13 seconds. So in this case refactoring wasn't extremely helpful and my code is much less efficent than the original. 


### My Recommendation: 
Ultimately, while there are a few promising green energy stocks, specifically, ENPH & RUN. I would advise our clients, Steve’s parents, to diversify their investment portfolio & be selective with their green energy investments. Their preferred investment of DQ may be a good buy in 2018 if the company is able to turn around their performance, but showing a 62% drop in market share should give them serious pause. 
