import pandas as pd
import numpy as np

# Function to load data
def load_data(inventory_path, sales_path, closing_invetory_path):
    """
    - inventory_path (str): location of inventory source file.
    - sales_path (str): location of sales source file.
    - margins_path (str): location of closing inventory source file, which contains margins
    - closing_inventory_path (str): location of closing inventory
    """
    inventory = pd.read_csv(inventory_path)
    sales = pd.read_csv(sales_path)
    closing_invetory = pd.read_excel(closing_invetory_path, skiprows=2)
    return inventory, sales, closing_invetory

# Function to preprocess and clean margins data
def preprocess_margins(closing_invetory_df: pd.DataFrame):
    """
    - margins_df (dataframe): resulted dataframes from load_data function.
    """
    margin_df = closing_invetory_df.copy()
    margin_df.rename(columns={'შიდა კოდი': 'code', 'რენტაბელობა დღგ-ს გარეშე (%) (Sum)': 'margin'}, inplace=True)
    margin_df = margin_df[['code', 'margin']]
    margin_df = margin_df[margin_df.margin != 'Error']
    margin_df.margin = margin_df.margin.astype('float')
    margin_df.margin = margin_df.margin * 100

    # get rid of duplicates
    margin_df = margin_df.groupby('code')['margin'].mean().reset_index(drop=False)
    
    return margin_df

# preprocess closing_inventory
def preprocess_closing_inventory(closing_inventory: pd.DataFrame):

    cls_inventory = closing_inventory.copy()
    
    # remove columns
    remove_columns = ['საწყობი', 'შიდა კოდი',
        'შტრიხკოდი', 'საქონელი', 'კატეგორია',
        'ტიპი', 'რაოდენობა (Sum)']
    cls_inventory = cls_inventory[remove_columns]
    
    # rename columns
    column_names = ['warehouse', 'code', 'sku', 'product_name', 'category', 'type', 'quantity']
    cls_inventory.columns = column_names
    
    # fill down warehouse
    cls_inventory.warehouse.ffill(inplace=True)
    
    # remove empty values
    cls_inventory = cls_inventory.loc[~cls_inventory.code.isna(), :].reset_index(drop=True)
    
    # change sku data type
    cls_inventory.sku = cls_inventory.sku.astype('str')
    
    # keep only necessary warehouses
    # filter warehouses
    warehouses_of_interest = ['1610000100 - პიქსელი საწყობი',
    '1610000200 - მარჯანიშვილი საწყობი',
    '1610000500 - ბათუმი საწყობი',
    '1610010100 - პიქსელი - ფილიალი 1',
    '1610011000 - ცენტრალური საწყობი (სანზონა)',
    '1610011400 - ისთ ფოინთი საწყობი',
    '1610020100 - მარჯანიშვილი - ფილიალი 2',
    '1610041100 - რუსთაველის - ფილიალი 8',
    '1610041500 - რუსთაველი 8 საწყობი',
    '1610050100 - ბათუმი მაღაზია',
    '1610070100 - თბილისი მოლი - ფილიალი 7',
    '1610071400 - თბილისი მოლი საწყობი',
    '1610080100 - ბათუმი XS - ფილიალი',
    '1610090100 - პეკინი',
    '1610100100 - ისთ ფოინთი - ფილიალი 10',
    '1610110100 - ყაზბეგი',
    '1610111400 - ყაზბეგი საწყობი']
    cls_inventory = cls_inventory[cls_inventory.warehouse.isin(warehouses_of_interest)]
    
    # remove unnecessary categories
    remove_categories = ['საშობაო', 'დეკორაცია-აქსესუარები', 'საზაფხულო',
       'სათამაშოები', 'აქსესუარები', 'სახის მოვლა',
       'ტანის მოვლა', 'ჭურჭელი', 'სამზარეულო',
       'ჩანთები', 'ჰიგიენა', 'საკანცელარიო', 'აბაზანა', 'ტექსტილი',
       'საოჯახო აქსესუარები', 'მობილურის აქსესუარები',
       'კოსმეტიკის აქსესუარები', 'თმის აქსესუარები', 'ტექნიკა',
       'ფიტნესი ', 'ჰაერის არომათერაპია', 'ელემენტები',
       'ზამთრის აქსესუარები', 'თმის მოვლა', 'ბიჟუტერია ',
       'შინაური ცხოველები', 'გაზაფხული-შემოდგომის აქსესუარები', 'ჩუსტები',
       'სასაჩუქრე ნაკრები', 'კომპიუტერის აქსესუარები',
       'პარფიუმერია', 'კოსმეტიკა']
    
    cls_inventory = cls_inventory[cls_inventory.category.isin(remove_categories)].reset_index(drop=True)

    return cls_inventory

# Function to preprocess inventory and sales data
def preprocess_inventory_sales(inventory, sales):
    """
    - inventory, sales (dataframe): resulted dataframes from load_data function
    """
    inventory = inventory[inventory.cogs != 0]
    inventory['date'] = pd.to_datetime(inventory['year'].astype(str) + '-' + inventory['month'].astype(str))
    sales['date'] = pd.to_datetime(sales.date)
    return inventory, sales

# Function to perform DSI analysis
def dsi_analysis(inventory, sales, inventory_with_dates):
    """
    - inventory, sales (dataframe): resulted dataframes from load_data function
    """
    # Implement the DSI analysis as per the provided code
    # This will involve separating closing inventory, calculating DSI, etc.
    # seperate closing inventory
    
    closing_inventory = inventory.copy().groupby(['code', 'product_name', 'category', 'type']).agg(
        {'quantity': 'sum'}
    ).reset_index(drop=False)
    # get min dates for each code
    min_dates = inventory_with_dates.groupby('code')['date'].min()
    # get max date
    max_dates = inventory_with_dates.groupby('code')['date'].max()
    # map with closing_inventory dataframe
    closing_inventory['open_date'] = closing_inventory['code'].map(min_dates)
    # closing_inventory.rename({'date': 'open_date'}, axis=1, inplace=True)
    closing_inventory.open_date = pd.to_datetime(closing_inventory.open_date)
    
    closing_inventory['date'] = closing_inventory['code'].map(max_dates)
    closing_inventory.rename({'date': 'closing_date'}, axis=1, inplace=True)
    closing_inventory.closing_date = pd.to_datetime(closing_inventory.closing_date)

    
    # change date to datetime
    sales.date = pd.to_datetime(sales.date)
    # group sales by code, aggregating cogs, profit, last sales date
    sales_grouped = sales.groupby('code')[['cogs', 'quantity', 'date']].agg(
        {
            'cogs': 'sum',
            'quantity': 'sum',
            'date': 'max'
        }
    ).reset_index(drop=False)
    # combine sales_grouped and closing_inventory
    closing_inventory = pd.merge(left=closing_inventory, right=sales_grouped, on='code', how='left')
    # calculate number of days product has been selling
    closing_inventory['days_being_sold'] = (closing_inventory.closing_date - closing_inventory.open_date).dt.days
    # closing_inventory.loc[closing_inventory.days_being_sold.isna(), 'days_being_sold'] = 365
    
    # calculate DSI
    closing_inventory['DSI'] = (closing_inventory.days_being_sold / closing_inventory.quantity_y) * closing_inventory.quantity_x
    return closing_inventory

# Function to perform ABC analysis
def abc_analysis(sales):
    # Implement the ABC analysis as per the provided code
    # group by code, aggregaste profit
    sales_abc = sales.groupby('code')['profit'].sum().reset_index(drop=False).sort_values('profit', ascending=False)
    # calculate running total
    sales_abc['running_total'] = sales_abc.profit.cumsum()
    # calculate percentage of running_total to total_sum
    sales_abc['perc'] = (sales_abc.running_total / sales_abc.profit.sum()) * 100
    # rank running_total
    sales_abc['rank'] = sales_abc.running_total.rank()
    # perc of rank
    sales_abc['rank_perc'] = (sales_abc['rank'] / sales_abc['rank'].max()) * 100
    # A category <= 20%, B - > 20% <= 60%, C - > 60%
    sales_abc['ABC'] = sales_abc.apply(lambda row: 'A' if row.perc <= 80 else
                    ('C' if row.perc >= 95 else 'B'), axis=1)
    
    return sales_abc

# Function to perform XYZ analysis
def xyz_analysis(sales):
    # Implement the XYZ analysis as per the provided code
    # xyz dataframe
    sales_xyz = sales[sales.quantity != 0].groupby('code')['quantity'].agg(['std', 'mean']).reset_index(drop=False)
    # fill na values in std
    sales_xyz['std'].fillna(999, inplace=True)
    # replace mean 0 with 1
    sales_xyz['mean'].replace(0, 1, inplace=True)
    # replace mean 0 with 1
    sales_xyz['mean'].replace(-1, 1, inplace=True)
    # replace std 0 with 999
    sales_xyz['std'].replace(0, 999, inplace=True)
    # Calculate coefficient of variance
    sales_xyz['CV'] = sales_xyz['std'] / sales_xyz['mean']
    # assign XYZ
    sales_xyz['XYZ'] = sales_xyz.apply(lambda row: 'X' if row['CV'] < 0.5 else
                                    ('Z' if row['CV'] > 1 else 'Y'), axis=1)
    
    return sales_xyz

# Function for calculating margins for each products
def margin_analysis(sales):
    # groupby code, aggregate by revenue and profit
    sales_margin = sales.groupby('code')[['revenue', 'profit']].sum().reset_index(drop=False)
    # drop 0 revenues
    sales_margin = sales_margin[sales_margin.revenue != 0].reset_index(drop=True)
    # calculate margin
    sales_margin['margin'] = (sales_margin['profit'] / sales_margin['revenue']) * 100
    
    return sales_margin

# Function for calculating opening_stock date and quantity for each product
def opening_stock_summary(inventory):
    # groupby min date for each code
    opening_stock = inventory.groupby('code')['date'].min().reset_index(drop=False)
    # merge to opening stock level
    opening_stock = pd.merge(left=opening_stock, right=inventory[['code', 'date', 'quantity']], on=['code', 'date'], how='left').groupby(['code', 'date'])['quantity'].sum().reset_index(drop=False)
    
    return opening_stock

# Function for calculating opening inventory for each products
def doh_analysis(inventory, opening_stock):
    """
    - inventory: whole inventory
    - opening_stock: from opening_stock_summary function
    """
    
    # list of codes in closing_inventory
    doh_df = inventory[inventory['date'] == inventory['date'].max()]
    # group by date and aggr by max date
    doh_df = doh_df.groupby('code')['date'].max().reset_index(drop=False)
    # merge opening stock's date
    doh_df = pd.merge(left=doh_df, right=opening_stock[['code', 'date']], on='code', how='left')
    # days on hand calculation
    doh_df['doh'] = (doh_df.date_x - doh_df.date_y).dt.days
    
    return doh_df

# Function to combine all analyses into a final dataframe
def combine_analyses(dsi, abc, xyz, margin, doh):
    """Combine analysis into single dataframe
    
    Keyword arguments:
    dsi, abc, xyz, margin, total_sales, closing_inventory, doh -- all of those are from functions
    Return: product_evaluation
    """
    
    # shows dsi and closing_inventory at the same time
    product_evaluation = dsi[['code', 'product_name', 'category', 'type', 'quantity_x', 'quantity_y','DSI']]
    
    # rename quantity to closing_inventory
    product_evaluation = product_evaluation.rename(columns={'quantity_x': 'closing_inventory', 'quantity_y': 'total_sales'})
    
    # merge with ABC
    product_evaluation = pd.merge(left=product_evaluation, right=abc[['code', 'ABC']], on='code', how='left')
    
    # merge with XYZ
    product_evaluation = pd.merge(left=product_evaluation, right=xyz[['code', 'XYZ']], on='code', how='left')
    
    # merge margin
    product_evaluation = pd.merge(left=product_evaluation, right=margin[['code', 'margin']], on='code', how='left')
        
    # merge with opening stock
    product_evaluation = pd.merge(left=product_evaluation, right=doh[['code', 'doh']], on='code', how='left')
    
    return product_evaluation

# summarize by dsi
def dsi_type_summary(product_evaluation):
    # group by types average DSI
    dsi_type = product_evaluation[(~product_evaluation.DSI.isna()) & (product_evaluation.DSI != np.inf) & (product_evaluation.DSI != -np.inf)].groupby('type')['DSI'].mean().reset_index(drop=False).sort_values('DSI').reset_index(drop=True)
    
    return dsi_type

# summarize by abc
def abc_type_summary(product_evaluation):
    # identify which category is most likely to be
    abc_type = product_evaluation[['type', 'ABC']].pivot_table(index='type', columns='ABC', aggfunc='size', fill_value=0).reset_index(drop=False)
    abc_type['total'] = abc_type['A'] + abc_type['B'] + abc_type['C']
    abc_type['A'] = (abc_type['A'] / abc_type['total']) * 100
    abc_type['B'] = (abc_type['B'] / abc_type['total']) * 100
    abc_type['C'] = (abc_type['C'] / abc_type['total']) * 100
    # expected category
    abc_type['expected_abc'] = abc_type[['A', 'B', 'C']].idxmax(axis=1)

    return abc_type

# summarize by xyz
def xyz_type_summary(product_evaluation):
    # identify which category is most likely to be
    xyz_type = product_evaluation[['type', 'XYZ']].pivot_table(index='type', columns='XYZ', aggfunc='size', fill_value=0).reset_index(drop=False)
    xyz_type['total'] = xyz_type['X'] + xyz_type['Y'] + xyz_type['Z']
    xyz_type['X'] = (xyz_type['X'] / xyz_type['total']) * 100
    xyz_type['Y'] = (xyz_type['Y'] / xyz_type['total']) * 100
    xyz_type['Z'] = (xyz_type['Z'] / xyz_type['total']) * 100
    # expected category
    xyz_type['expected_xyz'] = xyz_type[['X', 'Y', 'Z']].idxmax(axis=1)

    return xyz_type

# fill missing values in final dataframe and clean data
def fill_missing_values(product_evaluation, dsi_type, margin_closed_inventory, abc_type, xyz_type):
    """
    - product_evaluation: dataframe from combine_analysis function
    """
    final_file = product_evaluation
    
    # fill total sales with 0
    final_file.total_sales.fillna(0, inplace=True)
    
    # fill missing DSI
    dsi_mapping = dsi_type.set_index('type')['DSI'] # Create a mapping series from dsi_type
    final_file['temp_DSI'] = final_file['type'].map(dsi_mapping) # Use the mapping to create a temporary column for DSI values in product_evaluation
    final_file['DSI'] = final_file['DSI'].fillna(final_file['temp_DSI']) # Fill missing values in 'DSI' with those from 'temp_DSI'
    final_file.drop('temp_DSI', axis=1, inplace=True) # Drop the temporary column
    
    # questionable products that will be removed
    removed_codes = final_file[final_file.DSI.isna()]
    
    # drop the rest products, anevaluatable
    final_file.drop(removed_codes.index, axis=0, inplace=True)
    
    # fill missing ABC
    abc_mapping = abc_type.set_index('type')['expected_abc'] # Create a mapping series
    final_file['temp_ABC'] = final_file['type'].map(abc_mapping) # Use the mapping to create a temporary column
    final_file['ABC'] = final_file['ABC'].fillna(final_file['temp_ABC']) # Fill missing values
    final_file.drop('temp_ABC', axis=1, inplace=True) # Drop the temporary column
    
    # fill missing XYZ
    xyz_mapping = xyz_type.set_index('type')['expected_xyz'] # Create a mapping series
    final_file['temp_XYZ'] = final_file['type'].map(xyz_mapping) # Use the mapping to create a temporary column
    final_file['XYZ'] = final_file['XYZ'].fillna(final_file['temp_XYZ']) # Fill missing values
    final_file.drop('temp_XYZ', axis=1, inplace=True) # Drop the temporary column
    
    # if had been in stock for more than 6 months and more than 20
    final_file.loc[(final_file.doh > 182) & (final_file.closing_inventory > 20) & (final_file.total_sales <= 5), 'ABC'] = 'C'
    # if had been in stock for more than 6 months and more than 20
    final_file.loc[(final_file.doh > 182) & (final_file.closing_inventory > 20) & (final_file.total_sales <= 5), 'XYZ'] = 'Z'
    
    # fill missing margin
    margin_mapping = margin_closed_inventory.set_index('code')['margin'] # Create a mapping series
    final_file['temp_margin'] = final_file['code'].map(margin_mapping) # Use the mapping to create a temporary column
    final_file['margin'] = final_file['margin'].fillna(final_file['temp_margin']) # Fill missing values
    final_file.drop('temp_margin', axis=1, inplace=True) # Drop the temporary column
    
    # replace np.inf-s with 999 in DSI
    final_file.loc[final_file.DSI == np.inf, 'DSI'] = 999
    final_file.loc[final_file.DSI == -np.inf, 'DSI'] = 999
    
    # fill empty DOH-s with 30s
    final_file.loc[final_file.doh.isna(), 'doh'] = 30
    
    return final_file

# Main function to orchestrate the data loading, preprocessing, and analysis
def main():
    # loading files
    inventory_path = r'D:\excel db\yoyoso\inventory\inventory_clean.csv'
    sales_path = r'D:\excel db\yoyoso\sales\sales_cleaned.csv'
    closing_inventory_path = r'D:\excel db\yoyoso\inventory\closing_inv_margins\closing_inventory_margins.xlsx'
    
    # preprocessing
    inventory, sales, closing_inventory = load_data(inventory_path, sales_path, closing_inventory_path)
    margins_prec = preprocess_margins(closing_inventory)
    inventory_df, sales_df = preprocess_inventory_sales(inventory, sales)
    closing_inventory_df = preprocess_closing_inventory(closing_inventory)
    
    # Call other analysis functions and combine results
    # dsi_analysis, abc_analysis, xyz_analysis, combine_analyses, etc.
    dsi_df = dsi_analysis(closing_inventory_df, sales_df, inventory_df)
    abc_df = abc_analysis(sales_df)
    xyz_df = xyz_analysis(sales_df)
    calculate_margin = margin_analysis(sales)
    opening_stock = opening_stock_summary(inventory_df)
    doh_df = doh_analysis(inventory_df, opening_stock)
    
    product_evaluation = combine_analyses(dsi_df, abc_df, xyz_df, calculate_margin, doh_df)
    
    product_evaluation = product_evaluation.merge(right=margins_prec[['code']], on='code', how = 'right')
    
    """
    add types and the clean the data
    """
    dsi_by_types = dsi_type_summary(product_evaluation)
    abc_by_tyoes = abc_type_summary(product_evaluation)
    xyz_by_types = xyz_type_summary(product_evaluation)
    
    product_evaluation = fill_missing_values(product_evaluation, dsi_by_types, margins_prec, abc_by_tyoes, xyz_by_types)
    
    product_evaluation.to_csv("product_evaluation.csv", index=False)

if __name__ == "__main__":
    main()
