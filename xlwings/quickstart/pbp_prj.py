import xlwings as xw
import pandas as pd
from sqlalchemy import create_engine
import os
# from xlwings import Book, Range

def smmarize_sales():
    """
    Retrieve the account number and date ranges fro Excel sheet
    """
    # Make a connection to the calling excel file
    wb = xw.Book.caller()

    db_file = os.path.join(os.path.dirname(wb.fullname), 'pbp_proj.db')
    engine = create_engine(r"sqlite:///{}".format(db_file))

    # Retrieve the account number from the excel sheet as an int
    account = xw.Range('B2').options(numbers=int).value

    # Retrieve the account number and dates
    # account = xw.Range('B2').value
    start_date = xw.Range('D2').value
    end_date = xw.Range('F2').value

    # output the data just to make sure it all works
    # xw.Range('A5').value = account
    # xw.Range('A6').value = start_date
    # xw.Range('A7').value = end_date

    # clear existing data
    xw.Range('A5:F100').clear_contents()

    #create sql query
    # sql = 'SELECT * FROM sales WHERE account="{}" AND date BETWEEN "{}" AND "{}"'.format(account, start_date, end_date)
    sql = 'SELECT * FROM sales WHERE account="{}"'.format(account)

    # read query directly into the dataframe
    sales_data = pd.read_sql(sql, engine)

    # Analyze the data however we wanr
    summary = sales_data.groupby(["sku"])["quantity", "ext-price"].sum()

    total_sales = sales_data["ext-price"].sum()

    # output the results
    if summary.empty:
        xw.Range('A5').value = "No data for account {}".format(account)
    else:
        xw.Range('A5').options(index=True).value = summary
        xw.Range('E5').value = "Total Sales"
        xw.Range('F5').value = total_sales
    


