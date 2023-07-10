import pandas as pd     #Import Pandas for Datasets
import numpy as np      #Import Numpy for Dataset Filters
import yfinance as yf   #Import Yahoo Finance for Stock Info
import datetime         #Import Datetime for dates in Stock Info
import openpyxl         #Import OpenPYXL to save Excel Files

#Function to obtain 100% of Options Data for a Ticker
def options_chain(symbol):
    tk = yf.Ticker(symbol)
    # Expiration dates
    exps = tk.options

    # Get options for each expiration
    options = pd.DataFrame()
    for e in exps:
        opt = tk.option_chain(e)
        opt = pd.DataFrame().append(opt.calls).append(opt.puts)
        opt['expirationDate'] = e
        options = options.append(opt, ignore_index=True)

    # Bizarre error in yfinance that gives the wrong expiration date
    # Add 1 day to get the correct expiration date
    options['expirationDate'] = pd.to_datetime(options['expirationDate']) #+ datetime.timedelta(days = 1)
    options['dte'] = (options['expirationDate'] - datetime.datetime.today()).dt.days / 365
    
    # Boolean column if the option is a CALL
    options['CALL'] = options['contractSymbol'].str[4:].apply(
        lambda x: "C" in x)
    
    options[['bid', 'ask', 'strike']] = options[['bid', 'ask', 'strike']].apply(pd.to_numeric)
    options['mark'] = (options['bid'] + options['ask']) / 2 # Calculate the midpoint of the bid-ask
    
    # Drop unnecessary and meaningless columns
    options = options.drop(columns = ['contractSize', 'currency', 'change', 'percentChange', 'lastTradeDate', 'lastPrice','dte','inTheMoney','impliedVolatility','volume'])
    options = options.drop(columns = ['contractSymbol'])
    return options

#Options Data to Excel
SPYChain = options_chain("SPY")
filepath = r"C:\Users\Test\Desktop\Straddle Master\SPYChain.xlsx"
#print(SPYChain)

#Get Price of SPY Currently (Rounded to nearest Strike Price)
SPY = yf.Ticker("SPY")
data = SPY.history()
last_quote = round(data.tail(1)['Close'].iloc[0])
print("SPY",last_quote)

CallValue = np.where((SPYChain['CALL']==True) & (SPYChain['strike']== last_quote)) #Get List of Dates
ExpirationDates = SPYChain['expirationDate'].loc[CallValue]
SmolData= pd.DataFrame(["1","2","3","4","5","6","7"])

writer = pd.ExcelWriter('SPYChain.xlsx') #Open Writer of Excel File
SPYChain.to_excel(writer, sheet_name='SPY_Chain', index=False) #Add All Data to Sheet 1 
for date in ExpirationDates:
    DateFormatted = date.strftime("%Y-%m-%d")
    print(DateFormatted)
    Data = pd.DataFrame()
    Data['min'] = ''
    Data['span'] = ''
    Data['max'] = ''
    #Straddles surrounding middle
    for strike in range(last_quote-10,last_quote+10):
        CallValueLoc = np.where((SPYChain['CALL'] == True) & (SPYChain['strike'] == strike) & (SPYChain['expirationDate'] == date))
        PutValueLoc = np.where((SPYChain['CALL'] == False) & (SPYChain['strike'] == strike) & (SPYChain['expirationDate'] == date)) 
        Call = SPYChain.loc[CallValueLoc]
        Put = SPYChain.loc[PutValueLoc]
        Call = Call.reset_index(drop=True)
        Put = Put.reset_index(drop=True)
        if ((Call.size> 0) & (Put.size>0)):
            BreakEvenUp = Call.at[0,'mark'] + Put.at[0,'mark'] + strike
            BreakEvenDown = strike - Call.at[0,'mark'] - Put.at[0,'mark']
            Span = BreakEvenUp - BreakEvenDown
        Data = Data.append({'max' : BreakEvenUp,'span' : Span,'min' : BreakEvenDown}, ignore_index=True)
        Data = Data.append(Call)
        Data = Data.append(Put)
    
    Data = Data.append({'max' : "-",'span' : "-",'min' : "-",'strike' : "-",'bid' : "-",'ask' : "-",'openInterest' :"-",'expirationDate' :"-",'CALL' :"-",'mark' :"-"}, ignore_index=True)
    Data = Data.append({'max' : "@@@@@",'span' : "@@@@@",'min' : "@@@@@",'strike' : "@@@@@",'bid' : "@@@@@",'ask' : "@@@@@",'openInterest' :"@@@@@",'expirationDate' :"@@@@@",'CALL' :"@@@@@",'mark' :"@@@@@"}, ignore_index=True)
    Data = Data.append({'max' : "-",'span' : "-",'min' : "-",'strike' : "-",'bid' : "-",'ask' : "-",'openInterest' :"-",'expirationDate' :"-",'CALL' :"-",'mark' :"-"}, ignore_index=True)
    
    #Strangles surrounding middle
    for strike1 in range(last_quote,last_quote+10):
        for strike2 in range(last_quote-10,last_quote):
            CallValueLoc = np.where((SPYChain['CALL'] == True) & (SPYChain['strike'] == strike1) & (SPYChain['expirationDate'] == date))
            PutValueLoc = np.where((SPYChain['CALL'] == False) & (SPYChain['strike'] == strike2) & (SPYChain['expirationDate'] == date)) 
            Call = SPYChain.loc[CallValueLoc]
            Put = SPYChain.loc[PutValueLoc]
            Call = Call.reset_index(drop=True)
            Put = Put.reset_index(drop=True)
            if ((Call.size> 0) & (Put.size>0)):
                BreakEvenUp = Call.at[0,'mark'] + Put.at[0,'mark'] + strike1
                BreakEvenDown = strike2 - Call.at[0,'mark'] - Put.at[0,'mark']
                Span = BreakEvenUp - BreakEvenDown
            Data = Data.append({'max' : BreakEvenUp,'span' : Span,'min' : BreakEvenDown}, ignore_index=True)
            Data = Data.append(Call)
            Data = Data.append(Put)
    Data.to_excel(writer,sheet_name=DateFormatted, index=False)
writer.save()   



#for i in range(last_quote-5,last_quote+5):
#    StrikePrice = i
#    CallValue = np.where((SPYChain['CALL']==True) & (SPYChain['strike']== StrikePrice))
#    wb.create_sheet("Date")
#    print(SPYChain.loc[CallValue])
#wb.save(filepath)