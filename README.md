# TOS-RTD-core
framework for collecting real time stock price data from ThinkorSwim

### Context

There is a lot of buzz about institutional algorithmic trading on the stock market. There is also great interest in machine learning (here obviously). As an absolute beginner in ML with an interest in stock trading, I went looking for a dataset that I could use. 

While there are datasets available with daily closing prices, I couldn't find any publicly available with more granular data, much less anything close to real time.

I have a trading account with TD Ameritrade. Their main trading platform is called ThinkorSwim (TOS) and it has awesome capabilities. One feature is a Real Time Data (RTD) interface where Excel can pull price data from the platform at about a 3 second update rate. Using Visual Studio, I set up an Excel 'tunnel' to pull in data and then save it to a SQL Server DB. This is the method I used for that capture.

While the TOS RTD interface is one way (you cannot send trade orders to the platform), I can imagine a system that could deliver trading advice based on this near real time data.


