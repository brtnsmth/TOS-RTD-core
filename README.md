# TOS-RTD project

### Context

There is a lot of buzz about institutional algorithmic trading on the stock market. There is also great interest in machine learning (here obviously). As an absolute beginner in ML with an interest in stock trading, I went looking for a market dataset that I could use. 

While there are datasets available with daily closing prices (Yahoo finance), I couldn't find any publicly available with more granular data, much less anything close to having a real time data stream.

I have a trading account with TD Ameritrade. Their main desktop trading platform is called ThinkorSwim (TOS) and it has awesome capabilities. One feature is a Real Time Data (RTD) interface where Excel can pull price data from the platform at about a 3 second update rate. Using Visual Basic, I set up an Excel 'tunnel' to pull in data and then save it. 

This project is the core of that capture method. (My home system uses an SQL DB, but that is beyond basic core functionality....)

While the TOS RTD interface is one way (you cannot send trade orders to the platform), I can imagine a system that could deliver trading advice based on this near real time data. A real time multi-timescale momentum based price prediction tool maybe?



### Dataset Content

The dataset generated is 3 second interval price data using the futures trading windows - basically 24-7, from 6 PM EST Sunday to 6 PM Friday. A futures trading day starts at 6 PM on the day before a market trading day and ends at 5PM on the trading day. A new csv file is generated for each day.
 
The example group includes futures, stocks and ETFs. No attempt was made to limit specific captures to 'market open' time periods. 

There may also be a few exception events where data points are lost. A TOS refusal msg popped occassionally.



### Prerequisites

This project was built quick and dirty on an unspectacular Win10 machine. 

There are three common software packages involved -  
  - Visual Studio Community 2019
  - Excel - 2003 and 2010 versions were tested
  - ThinkorSwim - a free download for customers of TD Ameritrade. I STRONGLY encourage any trader to get a TDA account and this package. 
      -This project works on both the 'live trading' and the 'paper trading' TOS settings. 
      - Note - The TOS / TD EULAs limit their software to non-professional trader usage. 


### Project build

Create folders:
  - C:\temp\TOS Import core\deploy
  - C:\temp\TOS Import core\data

Visual Studio:
  - Start page - choose - Create a new project, VB, Windows, desktop, Windows Forms App
  - choose a project name and save location.

Go to form1 > view design - 
  - add a default label (Label1)
  - add a default timer (Timer1)
  - add a named timer (tmrExcelLoad)

Go to Project \ Properties \ References - 
  - add COM ref 'Microsoft Excel xx.x Object Library'. (xx.x is the version of Excel loaded on YOUR PC.)

Go to form1 > view code - 
- Copy & paste code in file 'TOS import core Form1.vb'

Build project / Save ALL


Start TOS. 
   - Note - Excel does not need to be open to run pgr.


In VS IDE - 
  - Press 'Start'

Normal events - 
  - Form1 shows.
  - Label text = 'Excel Load Timer started' (Default displays for about 10 seconds while Excel loads in background. Adjust for your PC load time)
  - Label text = "XL TOS.RTD connection tested OK" (displayed about 5 seconds)
  - Label text = "CSV data write success @ ... (write timestamp)


### Advanced (.exe) build:
Go to Project \ Properties \ Publish \ install mode
  - check 'This aplication is available offline..'
  - Select Updates
  - check 'The application should check for updates', 'Before application starts' and OK
change Publish Location to c:\temp\TOS import core\deploy\
Press 'Publish Now'

Go to c:\temp\TOS import core\deploy\ and run setup.exe.
OK through popups. Pgr should run. 

'TOS import core' shortcut should be available in Start menu to run pgr without IDE.













 

