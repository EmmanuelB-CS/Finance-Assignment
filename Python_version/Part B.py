import pandas as pd
import numpy as np
from scipy.optimize import minimize


# In this part of the code: I load the data from EURO STOXX Excel to compute the covariance 
# Matrix used in the computation of the portfolio standard deviation
###########################################################################################

excel_data_path = r"C:\Users\User\Desktop\Finance assignment\Python_version\EURO STOXX 50 Price + Index Data.xlsx"
df = pd.read_excel(excel_data_path, sheet_name='Price Data', index_col=0)

selected_stocks = ["ADIDAS (XET)", "ENEL", "KONINKLIJKE AHOLD DELHAIZE", "BBV.ARGENTARIA", "L AIR LQE.SC.ANYME. POUR L ETUDE ET L EPXTN.",
                       "AIRBUS", "ALLIANZ (XET)", "ANHEUSER-BUSCH INBEV", "ASML HOLDING", "AXA",
                       "BASF (XET)", "BAYER (XET)"]

selected_df = df[selected_stocks]
monthly_returns = selected_df.pct_change().dropna()
cov_matrix = monthly_returns.cov() # Computes the covariance matrix that will be used in the "optimization" function


# In this part of the code: I load the file I generated in part A in order to get the returns
# and standard deviation I will use as parameters in our optimization program
###########################################################################################
excel_file_path = r"C:\Users\User\Desktop\Finance assignment\csv_from_python_code\dataframes.xlsx"
sheet_name = 'Stocks'

stocks_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# I ask to the user whether or not he wants to consider a risk-free asset
user_choice = input("Do you want to add a risk-free asset ()? (yes/no): ").strip().lower()
if user_choice == 'yes':
    # Define the risk-free asset data
    rf_line_to_add = {'Stocks': 'Risk free', 
                     'Average Monthly Returns': 0.003, 
                     'Standard Deviation (Monthly)': 0.0,
                     'Annual Average Returns': 0.036, 
                     'Annual Standard Deviation': 0.0}
    
    # Append the risk-free asset line to the DataFrame
    stocks_data = stocks_data._append(rf_line_to_add, ignore_index=True)


returns = stocks_data[['Average Monthly Returns', 'Standard Deviation (Monthly)']] # Here, by using the append method I shall add the line of the risk free return
print(stocks_data)

# In this part of the code: I implement two functions that will compute both of the most 
# important values we need, the annual portfolio return and the annual porfolio stdv; they 
# are fundamental, mostly the second given it will be our function to minimize (please read
# the instructions below)
###########################################################################################
def annual_portfolio_return(weights, returns):
    """
    Calcule le rendement annuel d'un portefeuille en utilisant des poids donnés pour chaque actif.

    :param weights: Une liste des poids des actifs dans le portefeuille.
    :param returns: Un DataFrame des rendements des actifs avec les colonnes 'Average Monthly Returns'.
    :return: Le rendement annuel du portefeuille.
    """
    if len(weights) != len(returns):
        raise ValueError("Watch the number of weights! It does not equal the number of assets")

    portfolio_return = np.sum(returns['Average Monthly Returns'] * np.array(weights)) * 12  # Annualize
    return portfolio_return


def annual_portfolio_stddev(weights, cov_matrix):
    """
    Computes the portfolio's standard deviation using weight for each asset and a covariance matrix.

    :param weights: A list of the weights of the assets in the portfolio.
    :param cov_matrix: Covariance matrix of asset returns.
    :return: The standard deviation of the portfolio.
    """
    if len(weights) != cov_matrix.shape[0]:
        raise ValueError("The number of weights must match the number of assets in the covariance matrix")

    portfolio_variance = np.dot(weights, np.dot(cov_matrix, weights)) # equivalent of the following operation: XtAX (with X an array and A a square matrix)
    portfolio_stddev = np.sqrt(portfolio_variance)
    annual_portfolio_stddev = portfolio_stddev * np.sqrt(12)  # Annualize

    return annual_portfolio_stddev


# In this part of the code: here is the core thing, the function that will, for a set 
# of objective returns minimize the standard deviation of our portfolios and return the 
# optimal weights for as many objective returns we want!
######################################################################################
def optimize_portfolio(returns, cov_matrix, target_returns):
    """
    Optimize a portfolio with specified target returns.

    This function calculates an optimized portfolio with the objective of minimizing the portfolio's standard deviation
    while satisfying constraints on the target returns. The optimization is performed using the SciPy library's minimize
    function (so here can be the source of different results compared to the solver's version on Excel).

    :param returns: DataFrame with asset returns, including the 'Average Monthly Returns' column.
    :param cov_matrix: Covariance matrix of asset returns.
    :param target_returns: List of target annual returns to be achieved by the portfolio.
    :return: A list of dictionaries, each containing the portfolio's details for a specific target return.
             Each dictionary includes the 'Target Return', 'Portfolio Weight', 'Average Return', and 'Standard Deviation'.
    """
   
    results = []

    for target_return in target_returns:
        # Define optimization constraints for this target return
        constraints = [{'type': 'eq', 'fun': lambda x: np.sum(x) - 1}]
        constraints += [{'type': 'ineq', 'fun': lambda x: annual_portfolio_return(x, returns) - target_return}]

        initial_weights = [1 / len(selected_stocks)] * len(selected_stocks)  # Initial equal weights

        # Function to minimize
        fun_to_minimize = lambda x: annual_portfolio_stddev(x, cov_matrix)

        # Define bounds
        bounds = [(-1, 1) for _ in range(len(selected_stocks))]  # Weights should be between 0 and 1

        optimized_portfolio = minimize(fun_to_minimize, initial_weights, method='SLSQP', constraints=constraints, bounds=bounds)
        if optimized_portfolio.success:
            weights = optimized_portfolio.x
            return_, std = annual_portfolio_return(weights, returns), optimized_portfolio.fun
            results.append({'Target Return': target_return, 'Portfolio Weight': weights, 'Average Return': return_, 'Standard Deviation': std})

    return results



#######################################################################################
# Define target returns right here (please first read how np.arange works in the documentation)
target_returns = np.arange(0.10, 0.51, 0.10)  # 10% increments

# Call the function to optimize the portfolio
results = optimize_portfolio(returns, cov_matrix, target_returns)

for result in results:
    print("Target Return: {:.2%}, Portfolio Weight: {}, Average Return: {:.2%},"
      " Standard Deviation: {:.2%}".format(result['Target Return'],
                                           result['Portfolio Weight'],
                                           result['Average Return'],
                                           result['Standard Deviation']))

excel_writer = pd.ExcelWriter(r"C:\Users\User\Desktop\Finance assignment\csv_from_python_code\optimized_portfolios.xlsx", engine='xlsxwriter')

all_portfolios_df = pd.DataFrame()

for result in results:
    portfolio_df = pd.DataFrame({'Stocks': selected_stocks,
                                 'Portfolio Weight': result['Portfolio Weight'],
                                 'Average Return': result['Average Return'],
                                 'Standard Deviation': result['Standard Deviation'],
                                 'Target Return': result['Target Return']})
    
    current_target_return = result['Target Return']
    
    formatted_target_return = f"{current_target_return:.4%}"  # Format as percentage with 4 decimal places

    portfolio_df.to_excel(excel_writer, sheet_name=f'Target_Return_{formatted_target_return}', index=False)
    
    # Append the current portfolio DataFrame to the all_portfolios_df
    all_portfolios_df = pd.concat([all_portfolios_df, portfolio_df], ignore_index=True)

all_portfolios_df.to_excel(excel_writer, sheet_name='All Portfolios', index=False)

excel_writer._save()

