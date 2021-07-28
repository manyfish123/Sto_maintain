import xlrd
import pandas as pd

def readBuyerInfor():

    base_path = [
        'C:/Users/jimou/Downloads/Order.completed.20210628_20210728.xls',
        'C:/Users/jimou/Downloads/Order.completed.20210601_20210628.xls',
        'C:/Users/jimou/Downloads/Order.completed.20210501_20210531.xls'
        ]

    list_order_number = []
    list_buyer             = []
    list_buyer_payment     = []
    list_items_name_sale   = []
    list_items_amount_sale = []
    dic_buyer_payment = {}
    dic_buyer_orders = {}
    dic_test = {}
    total_payment = 0
    for item_path in base_path:

        book = xlrd.open_workbook(item_path)
        sheet1 = book.sheets()[0]
        nrows = sheet1.nrows
   
        for nu in range(1,nrows):

            order_nunber      = sheet1.cell(nu, 0).value
            order_buyer       = sheet1.cell(nu, 3).value
            order_time        = sheet1.cell(nu, 4).value
            order_price       = int(sheet1.cell(nu, 5).value.split('.')[0])

            order_item_name   = sheet1.cell(nu, 20).value
            order_item_amount = sheet1.cell(nu, 26).value

            if order_nunber not in list_order_number:
                 
                list_order_number.append(order_nunber)

                if order_buyer in dic_buyer_payment:

                    dic_buyer_payment[order_buyer] = dic_buyer_payment.get(order_buyer,0) + order_price
                    dic_buyer_orders[order_buyer] = dic_buyer_orders.get(order_buyer,0) + 1
                    
                else:
                    dic_buyer_payment[order_buyer] = order_price
                    dic_buyer_orders[order_buyer] = 1
                   
                
                total_payment += order_price
            if order_buyer not in list_buyer:
                list_buyer.append(order_buyer)

    dic_combine = {}
    for item in dic_buyer_payment:
        dic_combine[item] = [dic_buyer_orders[item],dic_buyer_payment[item]]
        #print('%s : %s : %s' %(item,dic_buyer_orders[item],dic_buyer_payment[item]))
    #print(dic_combine)
    dataFrame = pd.DataFrame(dic_combine,index=['購買次數','購買總金額'])
    dataFrame = pd.DataFrame.from_dict(dic_combine,columns=['購買次數','購買總金額'],orient='index').sort_values(by=['購買總金額'],ascending=False)
    print(dataFrame.sum())

readBuyerInfor()

