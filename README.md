# Value of risk calculation for portfolios in Dhaka Stock Exchange 

### Code Explanation
1. Data can be retrieved either from DSE website or from downloaded file. Variance Covariance method is implemented by default, other methods can also be incorporated by changing parameters.
2. VaR_U.Rmd is an R Notebook and generates a html page with the outputs, VaR_U.R is an R code file and writes outputs to an .xlsx file. VaR_Manual.R contains codes for calculating VaR manually (without utiliing built-in functions).
3. VaR.py is the python equivalent of VaR_Manual.R. The provided GoogleColab notebook executes the code.
4. However, it is recommended to use either VaR_U.Rmd or Var_U.R as they offer more functionalities.
5. We have to provide portfolio information in the form of port.xlsx, the file should have Instrument names in the first column, Market Values in the second and Cost Price in the third column. The column names should be 'Instrument', 'MKT_VAL' and 'Cost' respectively. 