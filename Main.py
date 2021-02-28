import pandas as pd
import difflib
import xlwt
from xlwt import Workbook

wb = Workbook()
sheet1 = wb.add_sheet('Sheet1')

df_imms =pd.read_excel('FILE DESTINATION')
df_cpt = pd.read_excel('FILE DESTINATION')

def cpt_deals():
    global deal_list
    global decision
    print("*TEST_NAMES 1 THROUGH 18*")
    client_name = int(input("Please choose from the following list:"))

    switcher = {
        1: "test_name 1",
        2: "test_name 2",
        3: "test_name 3",
        4: "test_name 4",
        5: "test_name 5",
        6: "test_name 6",
        7: "test_name 7",
        8: "test_name 8",
        9: "test_name 9",
        10: "test_name 10",
        11: "test_name 11",
        12: "test_name 12",
        13: "test_name 13",
        14: "test_name 14",
        15: "test_name 15",
        16: "test_name 16",
        17: "test_name 17",
        18: "test_name 18"

    }
    client = switcher.get(client_name, "Error: Please Try Again").split(', ')
    sub_client_df = df_cpt[df_cpt['ClientName'].str.strip().isin(client)]
    deal_list = sub_client_df['PortfolioName'].str.strip().tolist()
    print("Number of deals for this client: ", len(deal_list))
    sub_df = df_imms[df_imms['PortfolioName'].isin(deal_list)]
    print("IMMS Transactions: ", sub_df['IMMSTransactions'].sum())
    decision = str(input("Would you like to continue for another deal? \n y/n \n"))


def all_cpt_deals():
    client_list = ["*TEST_NAMES 1 THROUGH 18*"]

    for i in client_list:
        all_sub_client_df = df_cpt[df_cpt['ClientName'].isin([i])]
        all_deal_list = all_sub_client_df['PortfolioName'].tolist()
        all_sub_df = df_imms[df_imms['PortfolioName'].isin(all_deal_list)]
        print(i, ": ", all_sub_df['IMMSTransactions'].sum().tolist())


if __name__ == "__main__":
    task_type = str(input("Find IMMS Transactions for all clients? \n y/n \n"))
    if task_type == 'y':
        all_cpt_deals()
    elif task_type == 'n':
        decision = 'y'
        while decision == 'y':
            cpt_deals()

        print("Program Finished")
