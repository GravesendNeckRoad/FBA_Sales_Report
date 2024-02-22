# ____________________________________________________________________________________________________________________
# loading in necessary libraries/modules
# ____________________________________________________________________________________________________________________

# base
from datetime import date
import math
import os
import sys

# 3P
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns

# modules
from AutoFit_ColumnWidth import align_and_center
from DataBars_Creator import data_bars
from File_Readers import csv_reader, txt_reader, concat_df
from Style_Formatters import *
from Excel_Table_Creator import excel_table_creator

pd.set_option('display.width', None)

# ____________________________________________________________________________________________________________________
# reading in the raw data;
# ____________________________________________________________________________________________________________________
directory = r"C:\Users\taygu\Downloads\fba sales aggregator".replace("\\", "/")

sales = []
returns = []
instock = []
cogs = []

for file in os.listdir(directory):
    file_directory = os.path.join(directory, file)

    try:
        if file.endswith(".csv"):
            test = pd.read_csv(file_directory, nrows=1, encoding='latin1')

            if 'return-date' in test:
                returns.append(
                    csv_reader(
                        file_directory
                        , ['return-date', 'sku', 'quantity', 'product-name', 'reason', 'order-id']
                    )
                )

            elif 'afn-fulfillable-quantity' in test:
                instock.append(
                    csv_reader(
                        file_directory
                        , ['sku', 'afn-fulfillable-quantity', 'product-name']
                    )
                )

            elif 'VENDOR COST' in test:
                cogs.append(
                    csv_reader(
                        file_directory
                        , ['SKU', 'VENDOR COST']
                    )
                )

        elif file.endswith(".txt"):
            test = pd.read_csv(file_directory, sep='\t', nrows=1, encoding='latin1')
            if 'signature-confirmation-recommended ' in test:
                sales.append(
                    txt_reader(
                        file_directory
                        , [
                            'amazon-order-id'
                            , 'merchant-order-id'
                            , 'purchase-date'
                            , 'sku'
                            , 'quantity'
                            , 'product-name'
                            , 'item-price'
                        ]
                    )
                )

        else:
            continue

    except ValueError:
        print('Please review your files, one or more columns is missing.')
        sys.exit()

# terminating the program if any of the required files are missing
if any(not df for df in [sales, returns, instock, cogs]):
    print('ERROR: One or more of the required files necessary to run this report is missing. '
          'Please familiarize yourself with the documentation in the README and try again.')
    sys.exit()

# designating the product filter
brand_filter = input("Please enter the name of the brand/keyword you wish to run this report for.\n"
                     "If you wish to return all sales for the account(s) regardless of brand/keyword, "
                     "simply enter a single space, then press Enter:\n").strip()


# ____________________________________________________________________________________________________________________
# wrangling the data from the four sources;
# ____________________________________________________________________________________________________________________

# _____
# COGS;
# _____

try:
    cogs = concat_df(cogs)
    numeric_test_cogs = cogs['VENDOR COST'].mean()  # making sure the required values are numeric.
    # error will otherwise appear down in the in-stock block, since that's where you will be joining cogs to in-stock

except (pd.errors.DataError, ValueError, TypeError):
    print('Please review your `COGS` file(s). It appears you have some non-numeric data in them.')
    sys.exit()

if cogs.shape[0] == 0:
    print('It appears your COGS file is empty. Please ensure you are feeding a file with a SKU and UNIT COST column.')
    sys.exit()

# _____
# SALES;
# _____

try:
    sales = concat_df(sales) \
        .dropna() \
        .assign(purchase_date=lambda x: x['purchase-date'].str.split("T").str[0]) \
        .drop('purchase-date', axis=1) \
        .loc[lambda x:(x['product-name'].str.contains(f"{brand_filter}", case=False))] \
        .drop_duplicates() \
        .reset_index(drop=True)

    # creating a table that will contain date and revenue values that we will use in our table later
    sales_piv_w_dates = pd.pivot_table(
        sales
        , index='purchase_date'
        , values='item-price'
        , aggfunc='sum'
    ) \
        .reset_index() \
        .assign(
        MA7=lambda x: x['item-price'].rolling(7, min_periods=7).mean()
    )

    today = date.today()
    start_date = pd.to_datetime(sales_piv_w_dates['purchase_date']).min()
    end_date = pd.to_datetime(sales_piv_w_dates['purchase_date']).max()

    # if the report date range contains today's date, we will exclude today's date, since its not a complete sales day.
    # otherwise, when charting, it will deceptively appear as though there was a sharp drop in sales on the last day
    if str(today) == str(end_date.date()):
        end_date -= pd.Timedelta(days=1)

    # adding date range manually bc otherwise mlp and sns filter out 0's, which will improperly chart if you have days
    # with no sales in your data
    date_range = pd.date_range(
        start=start_date
        , end=end_date
        , freq='D'
    )

    df = pd.DataFrame({'Date': date_range})
    df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')  # getting rid of time from datetime, leaving just date

    sales_piv_w_dates = pd.merge(
        df,
        sales_piv_w_dates,
        left_on='Date',
        right_on='purchase_date',
        how='left'
    )

    # TODO
    # deleted this because i already added the time delta subtraction above. add back if it causes problems;
    # .loc[lambda x:(x['Date'] != end_date)]

    sales_piv_w_dates['item-price'] = sales_piv_w_dates['item-price'].fillna(0)

    sales_piv_no_dates = pd.pivot_table(
        sales
        , index=['sku', 'product-name']
        , values=['item-price', 'quantity']
        , aggfunc='sum'
    ) \
        .reset_index()

    total_sales_rev = sales['item-price'].sum().round(2)
    total_sales_qty = sales['quantity'].sum()

    sales_grandtotals = pd.DataFrame({
        'sku': ['Grand Totals']
        , 'quantity': [total_sales_qty]
        , 'item-price': [total_sales_rev]
    })

    sales_piv_no_dates = pd.concat(
        [sales_piv_no_dates, sales_grandtotals]
    ) \
        .sort_values(['item-price', 'quantity', 'product-name'], ascending=[0, 0, 1], na_position='first') \
        .reset_index(drop=True)

    # cleaning the column names and reordering for aesthetic appearance
    sales_piv_no_dates = sales_piv_no_dates[
        ['sku', 'product-name', 'item-price', 'quantity']
    ] \
        .rename({'item-price': 'gross revenue'}, axis=1)

except (KeyError, AttributeError):  # if the brand/keyword does not exist
    print('There are no sales for this keyword/brand name in your .txt sales file(s)')
    sys.exit()

except (pd.errors.DataError, ValueError, TypeError):
    print('Please review your `sales` files. It appears you have some non-numeric data in them.')
    sys.exit()

# _____
# IN STOCK;
# _____

try:
    instock = concat_df(instock) \
        .sort_values('afn-fulfillable-quantity', ascending=0, na_position='first') \
        .dropna() \
        .drop_duplicates() \
        .loc[lambda x:
            (x['afn-fulfillable-quantity'] > 0) &
            (x['product-name'].str.contains(f"{brand_filter}", case=False))] \
        .reset_index(drop=True)

    # now get the COGS into the in-stock table
    instock = pd.merge(
        instock,
        cogs,
        left_on='sku'
        , right_on='SKU'
        , how='left'
    ) \
        .reset_index()

    # if the cost ends up absent, the average of the costs will be taken, so as to not exclude anything.
    avg_vendor_cost = instock['VENDOR COST'].mean(skipna=True)

    instock['VENDOR COST'] = instock['VENDOR COST'].fillna(avg_vendor_cost)

    instock['total cost'] = instock['VENDOR COST'] * instock['afn-fulfillable-quantity']

    instock = instock[['sku', 'product-name', 'total cost', 'afn-fulfillable-quantity']]

    # ----- now in stock totals

    remaining_in_stock_qty = instock['afn-fulfillable-quantity'].sum()
    remaining_in_stock_value = instock['total cost'].sum()

    instock_grandtotals = pd.DataFrame({
        'sku': ['Grand Totals']
        , 'total cost': [remaining_in_stock_value]
        , 'afn-fulfillable-quantity': [remaining_in_stock_qty]
    })

    instock = pd.concat(
        [instock, instock_grandtotals]
        , ignore_index=True
    ) \
        .sort_values(['total cost', 'afn-fulfillable-quantity', 'product-name'], ascending=[0, 0, 1], na_position='first') \
        .reset_index(drop=True)

    instock = instock.rename({'afn-fulfillable-quantity': 'quantity'}, axis=1)

except (pd.errors.DataError, ValueError, TypeError):
    print('Please review your `in stock` file. It appears you have some non-numeric data in them.')
    sys.exit()

# _____
# RETURNS;
# _____

try:
    returns = concat_df(returns) \
        .assign(return_date=lambda x: x['return-date'].str.split("T").str[0]) \
        .loc[lambda x: x['product-name'].str.contains(f"{brand_filter}", case=False)] \
        .drop_duplicates() \
        .reset_index(drop=True)

    # TODO
    # creates error bc if there is only 1 sale then the date range is just 1 date - the day of sale, so this will break
    # returns = returns.loc[lambda x:
    # (x['return_date'].astype(str) >= str(start_date.date())) &
    # (x['return_date'].astype(str) <= str(end_date.date()))
    # ]  # test this to make sure it filters correctly, run a report with wack date range and make sure its ai

    # obtaining $ value of returns by order-id from sales
    returns = pd.merge(
        returns
        , sales[['amazon-order-id', 'sku', 'item-price']]
        , left_on=['order-id', 'sku']
        , right_on=['amazon-order-id', 'sku']
        , how='left'
    ) \
        .reset_index()

    # some $ values will inevitably be missing, so I will fill them with the average of whatever was found
    avg_return_value = returns.query('~`item-price`.isnull()')['item-price'].mean()
    returns['item-price'] = returns['item-price'].fillna(avg_return_value)

    total_return_value = (returns['item-price'] * returns['quantity']).sum()
    total_returns_qty = returns['quantity'].sum()

    returns_by_reason = pd.pivot_table(
        returns
        , index='reason'
        , values='quantity'
        , aggfunc='sum'
    ) \
        .reset_index() \
        .assign(
        pct_of_all_returns=lambda x: x['quantity'] / x['quantity'].sum()
    ) \
        .sort_values('pct_of_all_returns', ascending=0, na_position='first')

    returns_by_reason_totals = pd.DataFrame(
        {
            'quantity': [returns_by_reason['quantity'].sum()]
            , 'pct_of_all_returns': [returns_by_reason['pct_of_all_returns'].sum()]
        }
    )

    returns_by_reason = pd.concat([returns_by_reason, returns_by_reason_totals], ignore_index=True) \
        .sort_values(['reason', 'pct_of_all_returns'], ascending=[1, 0], na_position='first') \
        .rename({'reason': 'reason for return', 'pct_of_all_returns': 'percent of all returns'}, axis=1) \
        .reset_index(drop=True)

    # ---- now just regular returns

    returns = pd.pivot_table(
        returns
        , index=['sku', 'product-name']
        , values=['quantity', 'item-price']
        , aggfunc='sum'
    ) \
        .reset_index()

    returns_grandtotals = pd.DataFrame({
        'sku': ['Grand Totals']
        , 'quantity': [total_returns_qty]
        , 'item-price': [total_return_value]
    })

    returns = pd.concat(
        [returns, returns_grandtotals]
        , ignore_index=True
    ) \
        .sort_values(['quantity', 'product-name'], ascending=[0, 1], na_position='first') \
        .reset_index(drop=True) \
        .rename({'item-price': 'returned revenue'}, axis=1)
    # if you don't drop=true then the extra index col will appear

    returns = returns[['sku', 'product-name', 'returned revenue', 'quantity']]

except KeyError:  # if there are no returns for this keyword
    total_returns_qty = 0
    total_return_value = 0

    returns = pd.DataFrame({
        'sku': ['Grand Totals']
        , 'product-name': ['']
        , 'returned revenue': [total_return_value]
        , 'quantity': [total_returns_qty]
    })

    returns_by_reason = pd.DataFrame({
        'reason for return': ['No returns for this brand/timeframe.']
        , 'quantity': [0]
        , 'percent of all returns': [0]

    })

except (pd.errors.DataError, ValueError, TypeError):  # non numeric values in the numeric columns
    print('Please review your `returns` file(s). It appears you have some non-numeric data in them.')
    sys.exit()

# __________________________________________________________________________________________________________________
# writing cleaned and wrangled df's to excel files;
# __________________________________________________________________________________________________________________

# TODO
# add date logic here. if sales period only ~1 day then the title should just be for today
report_name = [f"{brand_filter.title()} Sales {start_date.strftime('%m-%d-%Y')} through {end_date.strftime('%m-%d-%Y')}"
               if len(brand_filter.title()) > 1 else
               f"Sales {start_date.strftime('%m-%d-%Y')} through {end_date.strftime('%m-%d-%Y')}"][0]

output_path = f"{directory}/{report_name}.xlsx"

with pd.ExcelWriter(output_path) as writer:
    pd.DataFrame().to_excel(writer, "cover page", index=False)  # blank sheet that will house our time series chart
    instock.to_excel(writer, "in stock", index=False)
    sales_piv_no_dates.to_excel(writer, "gross sales", index=False)
    returns.to_excel(writer, "returns", index=False)
    returns_by_reason.to_excel(writer, "returns by reason", index=False)
    # sales_piv_w_dates.to_excel(writer, "pivot table by day", index=False)  # for testing purposes only

# __________________________________________________________________________________________________________________
# open workbooks in openpyxl to format
# __________________________________________________________________________________________________________________

wb = xl.load_workbook(output_path)
ws_returns = wb['returns']
ws_instock = wb['in stock']
ws_sales = wb['gross sales']
ws_cover_page = wb['cover page']
ws_returns_by_reason = wb['returns by reason']

# __________________________________________________________________________________________________________________
# Create a time-series chart for the cover page, with some quick stats
# __________________________________________________________________________________________________________________

fig, ax = plt.subplots(figsize=(17, 7))

# MA7 has 7 first dates missing, so if the report is less than a week in length, you will get an error
# therefore, for data ranges less than a week, there will be no time series chart. just raw sales/returns/remaining
try:
    sns.lineplot(
        data=sales_piv_w_dates
        , x='Date'
        , y='MA7'
        , color='red'
        , linestyle='solid'
    )

    plt.stackplot(
        sales_piv_w_dates['Date']
        , sales_piv_w_dates['item-price']
    )

    plt.grid(
        True
        , color="grey"
        , linewidth=".1"
        , linestyle="-."
    )

    sns.despine()

    ax.set_xticks(sales_piv_w_dates['Date'][::7])

    plt.xticks(rotation=20, fontsize=12)

    plt.yticks(fontsize=13)

    plt.suptitle(
        f'{brand_filter.title()} Sales'
        , fontsize=23
    )

    plt.title(
        f'{start_date.strftime("%B %d, %Y")} - {end_date.strftime("%B %d, %Y")}'
        , fontsize=18
    )

    plt.text(
        .95
        , .5
        , f'____\n\n'
          f'Net Revenue: ${"{:,.2f}".format(total_sales_rev-total_return_value)}\n\n'
          f'Net Units Sold: {"{:,}".format(total_sales_qty-total_returns_qty)}\n'
          f'____\n\n'
          f'Avg Daily Revenue: ${"{:,.2f}".format(round((total_sales_rev-total_return_value)/len(date_range),2))} \n\n'
          f'Avg Daily Units Sold: {math.floor((total_sales_qty-total_returns_qty)/len(date_range))}-{math.ceil((total_sales_qty-total_returns_qty)/len(date_range))}\n'
          f'____\n\n'
          f'Overall Return Rate: {round((total_returns_qty / (total_sales_qty-total_returns_qty)) * 100,2)}%\n'
          f'____\n'
        , transform=plt.gca().transAxes
        , verticalalignment='bottom'
        , horizontalalignment='center'
        , bbox=dict(
            boxstyle='square'
            , facecolor='white'
            , alpha=0.5
        ), fontsize=12,
    )

    ax.set_ylabel('Daily Revenue', fontsize=13)

    ax.set_xlabel('')

    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, pos: f"${x:,.0f}"))

    # plt.show()  # for testing purposes

except:
    pass
    ws_cover_page.sheet_state = 'hidden'

# ___________________________________________________________________________________________________________________
# Visually format the workbook using openpyxl library functions
# ___________________________________________________________________________________________________________________

# todo
# should probably add this into the wrangling bricks, and the ws/wb import to the top of the sheet, so it can be
# more streamlined.

# if there are no returns, then we don't need to break down the customers reason for each return
if total_returns_qty == 0:
    ws_returns_by_reason.sheet_state = 'hidden'

to_make_bold = [
    (ws_instock, *['A2', 'C2', 'D2'])
    , (ws_sales, *['A2', 'C2', 'D2'])
    , (ws_returns, *['A2', 'C2','D2'])
    , (ws_returns_by_reason, *['A2', 'B2'])
]

to_highlight = [
    (ws_instock, 'FFFF00', *['C2', 'D2'])
    , (ws_sales, 'FFFF00', *['C2', 'D2'])
    , (ws_returns, 'FFFF00', *['C2', 'D2'])
    , (ws_returns_by_reason, 'FFFF00', 'B2')
]

inputs_for_databars = [
    (ws_returns, 3, '5e9bdd')
    , (ws_returns_by_reason, 2, '5e9bdd')
    , (ws_returns, 4, '5e9bdd')
    , (ws_instock, 3, '5e9bdd')
    , (ws_instock, 4, '5e9bdd')
    , (ws_sales, 3, '5e9bdd')
    , (ws_sales, 4, '5e9bdd')
]

[align_and_center(ws) for ws in [ws_returns, ws_instock, ws_sales, ws_returns_by_reason]]

[excel_table_creator(ws, table_name) for ws, table_name in
 zip([ws_returns, ws_instock, ws_sales, ws_returns_by_reason], ['Returns', 'InStock', 'Sales', 'SalesByReason'])]

# todo
# DISABLES DATABARS ON ALL SHEETS IF JUST THE IN-STOCK == 0. YOU NEED TO SEPARATE CLAUSES FOR EACH SHEET.
try:
    [data_bars(worksheet, column_number, color_hex) for worksheet, column_number, color_hex in inputs_for_databars
     if remaining_in_stock_qty > 0]
except:
    pass

[bold(worksheet, *cells) for worksheet, *cells in to_make_bold]

[highlighter(worksheet, color_hex, *cells) for worksheet, color_hex, *cells in to_highlight]

[thousands_sep_formatter(ws, cell) for ws, cell in zip([ws_instock, ws_returns, ws_sales], ['D2', 'D2', 'D2'])]

for ws in [ws_instock, ws_sales, ws_returns, ws_returns_by_reason]:
    col = 'B' if ws == ws_returns_by_reason else 'D'
    for cell in range(1, ws.max_row + 1):
        thousands_sep_formatter(ws, f'{col}{cell}')

[currency_formatter(ws, cell) for ws, cell in zip([ws_instock, ws_returns, ws_sales], ['C2', 'C2', 'C2'])]

for ws in [ws_instock, ws_sales, ws_returns]:
    for cell in range(1, ws.max_row + 1):
        currency_formatter(ws, f'C{cell}')

[percent_formatter(ws_returns_by_reason, f'C{cell}') for cell in range(1, ws_returns_by_reason.max_row + 1)]

[freeze_first_row(worksheet, 3) for worksheet in [ws_instock, ws_sales, ws_returns, ws_returns_by_reason]]

remove_gridlines(ws_cover_page)

# ___________________________________________________________________________________________________________________
# Add mpl chart to the cover page
# ___________________________________________________________________________________________________________________

img_path = f"{directory}/{brand_filter}_24ghp1.png"

plt.savefig(img_path)

img = xl.drawing.image.Image(img_path)

ws_cover_page.add_image(img, 'B3')

# ___________________________________________________________________________________________________________________
# Save and export
# ___________________________________________________________________________________________________________________

wb.save(output_path)


