# check if bdshare is installed, if not install
try:
    from bdshare import get_hist_data
except:
    !pip install bdshare
    from bdshare import get_hist_data

# import required modules
import pandas as pd
import numpy as np
from scipy.stats import norm
import plotly.offline as py
import plotly.graph_objs as go

# set the number format
pd.options.display.float_format = "{:.7f}".format

# Function for VaR calculation and output
def VaR(startdate, enddate, portfolio):

    # get the portfolio
    portfolio = pd.read_excel("port.xlsx")
    portfolio = portfolio.sort_values(["Instrument"])
    instruments = portfolio["Instrument"]
    weight = np.array(portfolio["MKT_VAL"]/portfolio["MKT_VAL"].sum())

    # extract historical data
    df = []

    for instruments in instruments: 
        stock_data = get_hist_data(startdate, enddate, instruments)
        df.append(stock_data)
        
    df = pd.concat(df)
    df = df.reset_index()
    df = df.sort_values(by = ["symbol", "date"])
    df = df[["date", "symbol", "ltp", "high", "low", "open", "close", "ycp", "trade", "value", "volume"]]
    df.columns = ["DATE", "TRADING.CODE", "LTP", "HIGH", "LOW", "OPENP", "CLOSEP", "YCP", "TRADE", "VALUE", "VOLUME"]
    cols = df.columns.drop(['DATE', 'TRADING.CODE'])
    df[cols] = df[cols].apply(pd.to_numeric, errors='coerce') 
    dat = df[["DATE", "TRADING.CODE", "CLOSEP"]]

    # reshape the data
    port_dat = dat.pivot(index = "DATE", columns = "TRADING.CODE", values = "CLOSEP").reset_index()

    # get the return series
    port_return = port_dat.iloc[1:,1:].diff()/port_dat.iloc[:-1,1:].values
    port_return = port_return.fillna(0)

    # variance covariance matrix of the return series
    varcov_mat = port_return.cov()

    # average return array of the portfolio 
    return_mat = np.array(port_return.mean())

    # average return of the portfolio 
    port_mean = weight.dot(return_mat)

    # porfolio standard deviation 
    port_var = weight.dot(varcov_mat).dot(weight.T)
    port_std = np.sqrt(port_var)

    # Value at risk for one day 
    VaR_pct = -port_mean + port_std * norm.ppf(.95)
    VaR = VaR_pct * portfolio["MKT_VAL"].sum()

    # get values for other time lengths
    days = np.array([5, 22, 22*2, 22*3])
    VaR_pct_next = VaR_pct * np.sqrt(days)
    VaR_next = VaR_pct_next * portfolio["MKT_VAL"].sum()

    # generate the output table
    days_col = np.append([1], days)
    var_pct = np.append([VaR_pct], VaR_pct_next)
    var = np.append([VaR], VaR_next)

    output = pd.DataFrame({'Days': days_col, 'VaR_PCT': var_pct, 'VaR': var})

    # VaR plot
    trace = [go.Scatter(x=output["Days"], y=output["VaR"], mode="lines+markers", marker=dict(size=15))]
    layout = go.Layout(title="Value at Risk Over Time", 
                xaxis=dict(title="Days"),
                yaxis=dict(title="Value at Risk"), 
                hovermode="closest")
    VaR_plot = go.Figure(data=trace, layout=layout)
    VaR_plot = py.iplot(VaR_plot)

    # return the output
    return(VaR_plot, output)
    
# get the VaR Plot
VaR_output = VaR("2020-07-01", "2020-12-31", "port.xlsx")

# get the VaR table
VaR_output[1]
