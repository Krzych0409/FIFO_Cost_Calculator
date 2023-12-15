import openpyxl
from colorama import init, Fore, Back
import pandas as pd


class Data:
    def get_launch_data(self, wb_source):
        self.ws_launch_data = wb_source['launch_data']

        # PURCHASES
        self.sheet_name_p = str(self.ws_launch_data['D11'].value)
        self.first_row_p = int(self.ws_launch_data['D12'].value)
        self.col_date_p = int(self.ws_launch_data['D13'].value)
        self.col_name_p = int(self.ws_launch_data['D14'].value)
        self.col_ppi_p = int(self.ws_launch_data['D15'].value)
        self.col_volume_p = int(self.ws_launch_data['D16'].value)
        self.col_unr_volume_p = int(self.ws_launch_data['D17'].value)
        # self.min_col_p = min(self.col_date_p, self.col_name_p, self.col_ppi_p, self.col_volume_p, self.col_unr_volume_p)
        # self.max_col_p = max(self.col_date_p, self.col_name_p, self.col_ppi_p, self.col_volume_p, self.col_unr_volume_p)

        # SALES
        self.sheet_name_s = str(self.ws_launch_data['G11'].value)
        self.first_row_s = int(self.ws_launch_data['G12'].value)
        self.col_date_s = int(self.ws_launch_data['G13'].value)
        self.col_name_s = int(self.ws_launch_data['G14'].value)
        self.col_ppi_s = int(self.ws_launch_data['G15'].value)
        self.col_volume_s = int(self.ws_launch_data['G16'].value)
        self.col_cogs_s = int(self.ws_launch_data['G17'].value)
        # self.min_col_s = min(self.col_date_s, self.col_name_s, self.col_ppi_s, self.col_volume_s, self.col_cogs_s)
        # self.max_col_s = max(self.col_date_s, self.col_name_s, self.col_ppi_s, self.col_volume_s, self.col_cogs_s)


class Transaction:
    def __init__(self, date, name, price_per_item, volume):
        self.date = date
        self.name = name
        self.price_per_item = price_per_item
        self.volume = volume

class Purchase(Transaction):
    all_purchases = []

    def __init__(self, date, name, price_per_item, volume, unrealized_volume):
        super().__init__(date, name, price_per_item, volume)
        self.unrealized_volume = unrealized_volume

        Purchase.all_purchases.append(self)

class Sale(Transaction):
    all_sales = []

    def __init__(self, date, name, price_per_item, volume, cost_of_goods_sold):
        super().__init__(date, name, price_per_item, volume)
        self.cost_of_goods_sold = cost_of_goods_sold

        Sale.all_sales.append(self)

class Portfolio:
    def __init__(self):
        self.transactions = []

    def add_transaction(self, transaction):
        self.transactions.append(transaction)


def main():
    excel_name = 'transactions.xlsx'

    # Colorama
    init(autoreset=True)

    wb_source = openpyxl.load_workbook(excel_name, data_only=False)

    data = Data()
    data.get_launch_data(wb_source)

    # Worksheet sales
    ws_s = wb_source[data.sheet_name_s]

    df_s = pd.read_excel(excel_name, sheet_name='sales', header=data.first_row_s-2)
    print(df_s)
    print()

    for row_s in df_s.itertuples():
        # Check cell == nan
        if pd.isna(row_s[data.col_name_s]) or not pd.isna(row_s[data.col_cogs_s]):
            continue

        print(row_s)

        df_p = pd.read_excel(excel_name, sheet_name='purchases', header=data.first_row_p-2)

        # Worksheet purchases
        ws_p = wb_source[data.sheet_name_p]

        required_volume = float(row_s[data.col_volume_s])
        missing_volume = required_volume
        found_volume = 0.0
        total_cost = 0.0

        for row_p in df_p.itertuples():
            if row_s[data.col_name_s].lower().strip() != row_p[data.col_name_p].lower().strip():
                continue
            if not row_p[data.col_unr_volume_p] > 0:
                continue
            if row_s[data.col_date_s] < row_p[data.col_date_p]:
                print(f'{Fore.YELLOW}The date of the sale transaction is later than the purchase transaction:'
                      f'\n{row_s[data.col_name_s]} : {row_s[data.col_volume_s]} : {row_s[data.col_ppi_s]} - still searching')
                continue

            new_unr_volume = float(row_p[data.col_unr_volume_p])
            print(f'{new_unr_volume} - must be float with ".0"')

            if row_p[data.col_unr_volume_p] >= missing_volume:
                new_unr_volume -= missing_volume  # new_unr_volume >= 0
                found_volume += missing_volume
                total_cost += float(row_p[data.col_ppi_p]) * missing_volume
                print(f'total_cost += {float(row_p[data.col_ppi_p])} * {missing_volume} = {float(row_p[data.col_ppi_p]) * missing_volume}')
                missing_volume -= missing_volume  # missing_volume = 0
            else:
                missing_volume -= new_unr_volume
                found_volume += new_unr_volume
                total_cost += float(row_p[data.col_ppi_p]) * new_unr_volume
                print(f'total_cost += {float(row_p[data.col_ppi_p])} * {new_unr_volume} = {float(row_p[data.col_ppi_p]) * new_unr_volume}')
                new_unr_volume -= new_unr_volume  # new_unr_volume = 0

            print(f'Unrealized volume (row_p = {row_p.Index + data.first_row_p}, column_p = {data.col_unr_volume_p}) = {new_unr_volume}')
            ws_p.cell(row=row_p.Index + data.first_row_p, column=data.col_unr_volume_p).value = new_unr_volume

            if missing_volume == 0 and found_volume == required_volume:
                print(f'End sale: {row_s[data.col_name_s]} - wolumen {row_s[data.col_volume_s]} - ppi {row_s[data.col_ppi_s]}')

                print(f'Total cost (row_s = {row_s.Index + data.first_row_s}, column_s = {data.col_cogs_s}) = {total_cost}')
                ws_s.cell(row=row_s.Index + data.first_row_s, column=data.col_cogs_s).value = total_cost

                try:
                    wb_source.save(excel_name)
                except PermissionError as err:
                    print(f'{Fore.RED}\nUnsaved data in excel file!\nMake sure the excel file is closed. Error details: \n{err}')

                print()
                print()
                break

        else:
            print(f'{Fore.RED}Not found all purchase transactions')



if __name__ == '__main__':
    main()




