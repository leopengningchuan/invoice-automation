# import packages
import pandas as pd


# define the function of getting the invoice information, set the default sales tax rate to 10%
def get_inv_info(
        inv_info, 
        inv_no, 
        sales_tax_rate = 0.1,
        ):
    
    # use the information for a specific invoice
    use_info = inv_info[inv_info['Invoice No.'] == inv_no].reset_index(drop = True)
    item_dict = {}
    
    # invoice information
    item_dict['CUSTOMER'] = use_info['Customer'].unique()[0]
    item_dict['CUSTOMER_ADDRESS1'] = use_info['Customer Address1'].unique()[0]
    item_dict['CUSTOMER_ADDRESS2'] = use_info['Customer Address2'].unique()[0]
    item_dict['INV_NO'] = use_info['Invoice No.'].unique()[0]
    item_dict['PAYMENT_TERMS'] = use_info['Payment Terms'].unique()[0]
    item_dict['DOC_DATE'] = str(pd.to_datetime(use_info['Invoice Date'].unique()[0]).date())
    item_dict['DUE_DATE'] = str(pd.to_datetime(use_info['Invoice Date'].unique()[0]).date() + pd.Timedelta(days=inv_info['Payment Terms'].unique()[0]))
    item_dict['SUB_AMOUNT'] = 0
    
    # product information for maximal of 15 items
    for i in range(1, 16):
        try:
            
            # for non empty items, get the information
            item_dict['ITEM' + str(i)] = use_info.loc[i-1, 'Item']
            item_dict['DETAIL' + str(i)] = use_info.loc[i-1, 'Detail']
            item_dict['UNITPRICE' + str(i)] = use_info.loc[i-1, 'Unit Price']
            item_dict['QUAN' + str(i)] = use_info.loc[i-1, 'Quantity']
            item_dict['AMT' + str(i)] = use_info.loc[i-1, 'Unit Price'] * use_info.loc[i-1, 'Quantity']
            
            # get the sum of subtotal
            item_dict['SUB_AMOUNT'] += item_dict['AMT' + str(i)]
            
        except:
            
            # for empty items, input empty information
            item_dict['ITEM' + str(i)] = ""
            item_dict['DETAIL' + str(i)] = ""
            item_dict['UNITPRICE' + str(i)] = ""
            item_dict['QUAN' + str(i)] = ""
            item_dict['AMT' + str(i)] = ""
    
    # get the tax and total amount
    item_dict['TAX_AMOUNT'] = round(item_dict['SUB_AMOUNT'] * sales_tax_rate, 2)
    item_dict['TOTAL_AMOUNT'] = item_dict['SUB_AMOUNT'] + item_dict['TAX_AMOUNT']
    
    # change all the unit price, amount to 000,000.00 format
    for k in item_dict.keys():
        if any(x in k for x in ('UNITPRICE', 'AMT', 'AMOUNT', 'QUAN', 'PAYMENT_TERMS')):
            try:
                if any(x in k for x in ('UNITPRICE', 'AMT', 'AMOUNT')):
                    item_dict[k] = '{0:,.2f}'.format(item_dict[k])
                else: 
                    item_dict[k] = '{0:,}'.format(item_dict[k])
            except:
                pass
            
    # return the item information
    return item_dict