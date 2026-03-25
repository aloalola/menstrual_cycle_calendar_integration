# Menstrual Cycle Calendar Integration
This is an Apps Script application to show past and predicted menstrual cycles in your Google Calendar. This makes living in sync with your cycle easier and allows you to plan according to your cycle phases. 

It's an easy integration to set up, even if you have no prior coding experience. Let me walk you through how that works step by step! :) 

# Set up the Cycle Tracker 
Open a new Google Sheets document from within your Google Drive and name it something like "Cycle Tracker". 
You then need to access Google Apps Script, which should show up in the toolbar under "Extensions" (sometimes you have to first [sign up](https://developers.google.com/apps-script) on Google Apps Script for it to show up.) 
If you click on that field in the toolbar, a new Apps Script document opens up. Name it somethinge like "Cycle Tracker" again. 

You now set up the basic tech, easy, right? 

You can now simply click on the code file I shared above and copy the entire code and paste it into your open Apps Script document. As we're working with on0pen, you do not have to specify which sheet you're referring to, it's simply using the ones that are open. 
Then press "run" within Apps Script to check whether you did that right. It should show something like "started" and "ended" on the bottom. 
Often, it now (or later) asks you to authorize the link you want to create to your calendar application. Simply agree by clickling the tiny greyed out text (instead of the large "back" button) and give Apps Script access to all the requested calendar features. 

Then simply switch to your Google sheet again. In the toolbar, a new tool should've appeared. This new functionality follows a three-step process: 
  1. Build the sheet: here, we're building the needed columns and set up the sheet.
  2. Build the cycles from your past cycles: here, we're using the data of your past cycles to build and especially *fix* the historical cycle events to remain untouched afterwards.
  3. Predict future cycles: here, we're predicting the new cycles based on the average cycle length of your past 6 months (a simple average).

But you may ask: how does my past cycle data get into my calendar? Great question! Between Step 1 and Step 2, you need to be a *freak in the sheets* yourself and put in your past cycle start dates (6 or more) in the Cycle_Log sheet. If you have a tracker app, check your latest first days of period and put them in in the format yyyy-mm-dd. In the other column ("notes"), you simply write "confirmed" in every cell. 
Good news! If you made it this far, you're basically unstoppable, because this data entry in Cycle_Log is the only place where you'll need to put in stuff manually. :) 

You're now good to move to Step 2. If you now click on the second button in the menu bar, it will rebuild the historical cycles and creates the events for that in your calendar. 
If this ran, you can already press the button for Step 3, which will predict your cycles for the next 6 months. 
And boom, set up is done. You can close the Apps Script and Sheet. 

# Using the Cycle Tracker 
If some time has passed and you now got your period, you have a new actual start date (not only predicted) for your menstrual cycle. Simply go into the excel sheet and add another date in the Cycle_log sheet, same as you did with your historical cycles. If you now press on the button for Step 3 again, your cycle will adapt, predictions will change, and your updated events will be in your calendar. 

# Some cool things to adapt 

## Cycle Name 
You can very easily change the name of your cycle tracker in your code (line 70 and 112). 

## Cycle Phase Length 
The current length of the cycle phases is fixed for menstruation, follicular phase, and ovulation, and then the remaining days "filled up" with luteal days up until the next cycle start. If you have a shorter cycle than 30 days, this might not be fully accurate. Maybe you know from your experience how long your menstruation usually lasts and when you roughly ovulate. You can simply adapt this in the code as well (line 72 to 75). It's super easy, trust me! 
