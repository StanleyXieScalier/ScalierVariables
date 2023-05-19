import numpy as np
import pandas as pd
from random import randrange
import warnings
import xlsxwriter
import locale
warnings.simplefilter("ignore")

# Read the Excel file
# Will be made to accomodate all reports later

file_path_1 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Placement report.xlsx'
file_path_2 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Campaign report.csv'
file_path_3 = 'E:/Scalier/MVP/Sample Data/BusinessReport-09-05-2023.csv'  #This is for 1 month here
file_path_4 = 'E:/Scalier/MVP/Sample Data/Jungle Scout Rank Tracker.xlsx'
file_path_5 = 'E:/Scalier/MVP/Sample Data/Keywords_1684403630.xlsx'
file_path_6 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Advertised product report.xlsx'
file_path_7 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Budget report.csv'
file_path_8 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Gross and invalid traffic report.xlsx'
file_path_9 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Performance over time report.xlsx'
file_path_10 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Purchased product report.xlsx'
file_path_11 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Search term impression share report.csv'
file_path_12 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Search term report.xlsx'
file_path_13 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Suchbegriff Bericht.xlsx'
file_path_14 = 'E:/Scalier/MVP/Sample Data/Sponsored Products Targeting report.xlsx'
file_path_15 = 'E:/Scalier/MVP/Sample Data/helium10-kt-B07HHGR987-2023-05-19.csv'
file_path_16 = 'E:/Scalier/MVP/Sample Data/Helium market tracker.xlsx'




def write_excel_to_number(file_path):
    writer = pd.ExcelWriter(file_path,
                            engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_numbers': True}})

#write_excel_to_number(file_path_14)


df_1 = pd.read_excel(file_path_1)
df_2 = pd.read_csv(file_path_2)
df_3 = pd.read_csv(file_path_3)
df_4 = pd.read_excel(file_path_4)
df_5 = pd.read_excel(file_path_5)
df_6 = pd.read_excel(file_path_6)
df_7 = pd.read_csv(file_path_7)
df_8 = pd.read_excel(file_path_8)
df_9 = pd.read_excel(file_path_9)
df_10 = pd.read_excel(file_path_10)
df_11 = pd.read_csv(file_path_11)
df_12 = pd.read_excel(file_path_12)
df_13 = pd.read_excel(file_path_13)
df_14 = pd.read_excel(file_path_14)
df_15 = pd.read_csv(file_path_15)
df_16 = pd.read_excel(file_path_16)


# Print the updated DataFrame
#print(df_3)


Product_name = 'Kordel gewachst'
Asin_name = 'B07HHGR987'


def check_column_type(data_frame, column_name):
    column_data = data_frame[column_name]

    if np.issubdtype(column_data.dtype, np.number):
        return 'Number'
    elif np.issubdtype(column_data.dtype, np.object):
        return 'String'
    else:
        return 'Unknown'
    

def convert_string_to_numeric(data_frame, column_name):
    #for column in data_frame.columns:
    if data_frame[column_name].dtype == np.object:
        try:
            data_frame[column_name] = pd.to_numeric(data_frame[column_name])
        except ValueError:
            print(f"Unable to convert column '{column_name}' to numeric.")



#Need to throw the exceptions (the ones that are not strings) here? 
def convert_comma_numbers_to_float(data_frame, column_name):
    if data_frame[column_name].str.contains('€', case=False).any(): #or .all()
        data_frame[column_name] = data_frame[column_name].str.replace(',', '').str.replace('€', '').astype(float)
    else:
        data_frame[column_name] = data_frame[column_name].str.replace(',', '').astype(float)



def dual_condition_new_df(data_frame, column1, column2, column3, A, B, title):

    # Iterate over the rows of the DataFrame
    new_data_frame = pd.DataFrame(columns=[title])

    # Iterate over the rows of the DataFrame
    for index, row in data_frame.iterrows():
        if A in str(row[column1]):
            position = row.name
            if B in str(data_frame.at[position, column2]):
                new_data_frame.loc[position, title] = data_frame.at[position, column3]


    # Add the new column to the DataFrame
    #data_frame['NewColumn'] = new_column

    return new_data_frame



def sum_columns(df, column1, column2):
    # Read the df
    data_frame = df

    # Sum the two columns
    sum_result = data_frame[column1] + data_frame[column2]

    return sum_result


#print(sum_columns(df_14, '7 Day Other SKU Sales', '7 Day Advertised SKU Sales'))


def sub_columns(df, column1, column2):
    # Read the df
    data_frame = df

    # minus the two columns
    sub_result = data_frame[column1] - data_frame[column2]

    return sub_result



def prod_columns(df, column1, column2):
    # Read the df
    data_frame = df

    # multiply the two columns
    prod_result = data_frame[column1] - data_frame[column2]

    return prod_result



def div_columns(df, column1, column2):
    # Read the df
    data_frame = df

    # divide the two columns
    div_result = data_frame[column1] / data_frame[column2]

    return div_result

#print(div_columns(df_14, '7 Day Other SKU Sales', '7 Day Advertised SKU Sales'))
#print(df_3['Page Views - Total'])
#print(df_3['Sessions - Total'])
result = check_column_type(df_3, 'Page Views - Total')
#print(result)
#convert_string_to_numeric(df_3, 'Page Views - Total')
#convert_comma_numbers_to_float(df_3, 'Page Views - Total')
result_2 = check_column_type(df_3, 'Page Views - Total')
#print(result_2)
#print(df_3['Page Views - Total'])
#print(div_columns(df_3, 'Sessions - Total', 'Page Views - Total'))



def Aggregate_asin_perf_mkt_metric_Calc(df):

    #can do Asin too
    #Impression
    #df.loc[df[Product_name] == 'Title', 'Sessions - Total']
    convert_comma_numbers_to_float(df, 'Sessions - Total') #Can be refined or put into a previous calc method
    Sessions = df.loc[df['(Parent) ASIN'].str.contains(Asin_name, case=False), 'Sessions - Total']
    convert_comma_numbers_to_float(df, 'Page Views - Total')
    CTR = div_columns(df, 'Sessions - Total', 'Page Views - Total') #Sessions/Views #Be careful with data type #Keep data with only chose product
    #visits/views?
    CR = div_columns(df, 'Sessions - Total', 'Units ordered') #orders/sessions #clean the NA,
    sessoin_per_order = div_columns(df, 'Units ordered', 'Sessions - Total') #Clicks_per_Order
    
    return Sessions, CTR, CR, sessoin_per_order

#Sessions, CTR, CR, sp_o=Aggregate_asin_perf_mkt_metric_Calc(df_3)
#print(sp_o)



def Aggregate_asin_perf_econ_perf_total_Calc(df, df_b, df_c):


    ordered_product_sales = df.loc[df['(Parent) ASIN'].str.contains(Asin_name, case=False), 'Ordered product sales']
    orders = df.loc[df['(Parent) ASIN'].str.contains(Asin_name, case=False), 'Units ordered']
    products_per_order = div_columns(df, 'Units ordered', 'Total order items') #makes it only for this product
    convert_comma_numbers_to_float(df_b, 'Spend')
    Actual_Acos_E = df_b.loc[df_b['Portfolio name'].str.contains(Product_name, case=False), 'Spend'].sum() #If we sum or not? #Or Campaign name?
    #Actual_Acos_percent  #could't find the data 
    Max_Acos_E = 1000  #this will be given
    Max_Acos_percent = "{0:.0%}".format(1) #will be given
    Delta_Actual_vs_Max_ACOS_E = Max_Acos_E - Actual_Acos_E   #check equation
    #Delta Actual vs Max ACOS %  #could't find the data 
    Num_of_KWs_in_top50 = df_c.loc[df_c['ASIN'].str.contains('B07FYQYF6B', case=False), 'Top 10'] #Wrong ASIN, right one not available. Change in future
    Search_Volume_KWs_top50 = df_c.loc[df_c['ASIN'].str.contains('B07FYQYF6B', case=False), 'Top 50 Search Volume'] #Wrong ASIN, right one not available. Change in future

    return ordered_product_sales, orders, products_per_order, Actual_Acos_E, Max_Acos_E, \
        Max_Acos_percent, Delta_Actual_vs_Max_ACOS_E, Num_of_KWs_in_top50, Search_Volume_KWs_top50

#op_s, o, ppo, aae, mae, map, dama, nkt5, svkt5 = Aggregate_asin_perf_econ_perf_total_Calc(df_3, df_2, df_4)
#print(nkt5 - svkt5) #extra 0 in front is just the index #Some with index some not?



#Seperate method or merge with the above?
def Aggregate_asin_perf_econ_perf_per_order_Calc(df, df_b):
    

    Product_Sales_pO_net = div_columns(df, 'Units ordered', 'Total order items').sum() #makes it only for this product
    convert_comma_numbers_to_float(df_b, 'Spend')
    ACOS_pO = df_b.loc[df_b['Portfolio name'].str.contains(Product_name, case=False), 'Spend'].sum() / \
    df.loc[df['(Parent) ASIN'].str.contains(Asin_name, case=False), 'Units ordered']
    Max_ACOS_pO_E = 100 #pre-set # change later
    Max_ACOS_pO_percent = "{0:.0%}".format(0.5) #pre-set #change later
    Delta_Actual_vs_Max_ACOS_pO_E = Max_ACOS_pO_E - ACOS_pO
    #Delta Actual vs Max ACOS pO %  ##could't find the data
    

    return Product_Sales_pO_net, ACOS_pO, Max_ACOS_pO_E, Max_ACOS_pO_percent, Delta_Actual_vs_Max_ACOS_pO_E

#psn, apo, mapoe, mapop, damape = Aggregate_asin_perf_econ_perf_per_order_Calc(df_3, df_2)
#print(damape)



ad_exp_budget = 1000
def Ad_group_exp_Camp_level(df, ad_exp_budget):


    budget = ad_exp_budget  #pre defined
    convert_comma_numbers_to_float(df, 'Spend')
    Total_spend = df.loc[df['Portfolio name'].str.contains(Product_name, case=False), 'Spend'].sum()
    budget_left = ad_exp_budget - Total_spend  #The spend should be monthly spend

    campaign_level_cpc = 300  #Will be set by us
    avg_cpc = df.loc[df['Portfolio name'].str.contains(Product_name, case=False), 'Spend'] / \
    df.loc[df['Portfolio name'].str.contains(Product_name, case=False), 'Clicks']
    total_impressions = df.loc[df['Portfolio name'].str.contains(Product_name, case=False), 'Impressions'].sum()
    total_clicks = df.loc[df['Portfolio name'].str.contains(Product_name, case=False), 'Clicks'].sum()
    avg_CTR = total_clicks / total_impressions  #Views or impressions as we used here? #For time period. Time period is already factored into the report. 
    avg_CR = df.loc[df['Portfolio name'].str.contains(Product_name, case=False), '7 Day Total Orders (#)'].sum() / total_clicks
    convert_comma_numbers_to_float(df, '7 Day Total Sales')
    seven_day_total_sales = df.loc[df['Portfolio name'].str.contains(Product_name, case=False), '7 Day Total Sales'].sum()  # do we need sum here?
    seven_day_total_orders = df.loc[df['Portfolio name'].str.contains(Product_name, case=False), '7 Day Total Orders (#)'].sum() # do we need sum here? or in many other places?
    sales_per_order = seven_day_total_sales / seven_day_total_orders
    #Acos in this report and seven day conversion rate? 


    return budget, Total_spend, budget_left, campaign_level_cpc, avg_cpc, total_impressions, total_clicks, \
            avg_CTR, avg_CR, seven_day_total_sales, seven_day_total_orders, sales_per_order

#b, ts, bl, clc, ac, ti, tc, act, acr, sdts, sdto, spo = Ad_group_exp_Camp_level(df_11, ad_exp_budget)
#print(tc)



def Ad_group_exp_placement_first_p(df):

    #Make sure the type is not object
    first_p_cpc_bid = 100  #pre set
    C1 = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Cost Per Click (CPC)', Product_name,'First page', 'avg_cpc') #take avg needs a new name
    avg_cpc = C1.sum().astype(float) / C1.count().astype(float) #Be sure to add as float to others after .sum()
    spend = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Spend', Product_name,'First page', 'spend').sum().astype(float)
    total_impressions = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Impressions', Product_name,'First page', 'total_impressions').sum().astype(float)
    total_clicks = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Clicks', Product_name,'First page', 'total_clicks').sum().astype(float)
    avg_CTR = total_clicks['total_clicks'] / total_impressions['total_impressions'] #declare the column name so result as just numbers
    seven_day_total_sales = dual_condition_new_df(df, 'Campaign Name', 'Placement', '7 Day Total Sales', Product_name,'First page', '7 Day Total Sales').sum().astype(float)
    seven_day_total_orders = dual_condition_new_df(df, 'Campaign Name', 'Placement', '7 Day Total Orders (#)', Product_name,'First page', '7 Day Total Orders').sum().astype(float)
    avg_CR = seven_day_total_orders['7 Day Total Orders'] / total_clicks['total_clicks']
    sales_per_order = seven_day_total_sales['7 Day Total Sales'] / seven_day_total_orders['7 Day Total Orders']


    return first_p_cpc_bid, avg_cpc, spend, total_impressions, total_clicks, avg_CTR, avg_CR, seven_day_total_sales, seven_day_total_orders, sales_per_order

#f, ac, s, ti, tc, act, acr, sdts, sdto, spo = Ad_group_exp_placement_first_p(df_1)
#print(spo)


def Ad_group_exp_placement_product_p(df):

    #Make sure the type is not object
    product_p_cpc_bid = 100  #pre set
    C1 = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Cost Per Click (CPC)', Product_name,'Product page', 'avg_cpc') #take avg needs a new name
    avg_cpc = C1.sum().astype(float) / C1.count().astype(float) #Be sure to add as float to others after .sum()
    spend = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Spend', Product_name,'Product page', 'spend').sum().astype(float)
    total_impressions = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Impressions', Product_name,'Product page', 'total_impressions').sum().astype(float)
    total_clicks = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Clicks', Product_name,'Product page', 'total_clicks').sum().astype(float)
    avg_CTR = total_clicks['total_clicks'] / total_impressions['total_impressions'] #declare the column name so result as just numbers
    seven_day_total_sales = dual_condition_new_df(df, 'Campaign Name', 'Placement', '7 Day Total Sales', Product_name,'Product page', '7 Day Total Sales').sum().astype(float)
    seven_day_total_orders = dual_condition_new_df(df, 'Campaign Name', 'Placement', '7 Day Total Orders (#)', Product_name,'Product page', '7 Day Total Orders').sum().astype(float)
    avg_CR = seven_day_total_orders['7 Day Total Orders'] / total_clicks['total_clicks']
    sales_per_order = seven_day_total_sales['7 Day Total Sales'] / seven_day_total_orders['7 Day Total Orders']


    return product_p_cpc_bid, avg_cpc, spend, total_impressions, total_clicks, avg_CTR, avg_CR, seven_day_total_sales, seven_day_total_orders, sales_per_order
    #make sure the returns are the same type

#p, ac, s, ti, tc, act, acr, sdts, sdto, spo = Ad_group_exp_placement_product_p(df_1)
#print(spo)


def Ad_group_exp_placement_rest(df):

    #Make sure the type is not object
    rest_cpc_bid = 100  #pre set
    C1 = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Cost Per Click (CPC)', Product_name,'Rest of search', 'avg_cpc') #take avg needs a new name
    avg_cpc = C1.sum().astype(float) / C1.count().astype(float) #Be sure to add as float to others after .sum()
    spend = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Spend', Product_name,'Rest of search', 'spend').sum().astype(float)
    total_impressions = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Impressions', Product_name,'Rest of search', 'total_impressions').sum().astype(float)
    total_clicks = dual_condition_new_df(df, 'Campaign Name', 'Placement', 'Clicks', Product_name,'Rest of search', 'total_clicks').sum().astype(float)
    avg_CTR = total_clicks['total_clicks'] / total_impressions['total_impressions'] #declare the column name so result as just numbers
    seven_day_total_sales = dual_condition_new_df(df, 'Campaign Name', 'Placement', '7 Day Total Sales', Product_name,'Rest of search', '7 Day Total Sales').sum().astype(float)
    seven_day_total_orders = dual_condition_new_df(df, 'Campaign Name', 'Placement', '7 Day Total Orders (#)', Product_name,'Rest of search', '7 Day Total Orders').sum().astype(float)
    avg_CR = seven_day_total_orders['7 Day Total Orders'] / total_clicks['total_clicks']
    sales_per_order = seven_day_total_sales['7 Day Total Sales'] / seven_day_total_orders['7 Day Total Orders']


    return rest_cpc_bid, avg_cpc, spend, total_impressions, total_clicks, avg_CTR, avg_CR, seven_day_total_sales, seven_day_total_orders, sales_per_order
    #make sure the returns are the same type

#r, ac, s, ti, tc, act, acr, sdts, sdto, spo = Ad_group_exp_placement_product_p(df_1)
#print(spo)



def Ad_group_exp_placement_lose_match(df, df_b):

    cpc_bid = 100  #set by us
    #avg_cpc
    spend = dual_condition_new_df(df_b, 'Campaign Name', 'Placement', 'Spend', Product_name,'Rest of search', 'Spend').sum().astype(float)
    #avg_search_term_impression_rank
    #search_term_impression_share
    #search_volume_kw_wc
    #num_kw_w_impression
    #num_kw_w_click
    #num_kw_w_sale
    #total_impressions
    #total_clicks
    #avg_CTR
    #avg_CR
    #clicks_per_order
    #seven_day_ad_sku_sales
    #seven_day_ad_sku_units
    #seven_day_other_sku_sales
    #seven_day_ad_sku_units
    #seven_day_total_sales
    #seven_day_total_orders
    #sales_per_total_orders
    #Acos_per_total_orders


    return 

#Ad_group_exp_placement_lose_match(df_11, df_14)
#print()



#######################################################

##Appending the calculated data to new sheets##
##Can be here below##
##Needs to be a different file from the calculation fiel##
##Making the methods into a class##


