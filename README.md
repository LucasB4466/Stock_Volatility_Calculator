# [Stock_Volatility_Calculator] (https://raw.githubusercontent.com/HotBreakfast/Stock_Volatility_Calculator/master/VBA_Code_HistoricalVolatility_Calculator "Calculator")
#-------
<P/>This workbook will ask for stock tickers when it opens. It will then pull the attributes of the tickers and ask for a date range</P>
<P/>to measure the historical volatility of the stock tickers. The end result displays a correlation matrix, historical volatility,</P> <P/> and attributes of the stock. You can copy this code and put it into a module. </P>
#-------


##Add the following worksheet names:

	Test
	Volatility
	Data
	Matrix

In the workbook module add the following:

Private Sub Workbook_Open()
RunMe
End Sub

Close workbook and open to run the macro.
#-------



##<P/>This workbook will ask you for stock tickers when it opens. It will then pull the attributes  </P>

<P/>The end result displays a correlation matrix, historical volatility, and attributes of the stock.  </P>
<P/>You can copy [this code] (https://raw.githubusercontent.com/HotBreakfast/Stock_Volatility_Calculator/master/VBA_Code_HistoricalVolatility_Calculator "this code") and put it into a module. </P>


#<P/>In the workbook module add the following:

##<P/>Private Sub Workbook_Open()
##<P/>RunMe
##<P/>End Sub

<P/>Close workbook and open to run the macro.
<P/>VBA code to get the historical volatility of stocks over a date range. The end result shows the volatility, 
a correlation matrix, and attributes of the stock tickers
