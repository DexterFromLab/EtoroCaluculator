import openpyxl

ETORO_ACCOUNT_STATEMENT_FILE_LOCATION = './etoro-account-statement-1-1-2022-12-31-2022.xlsx'
CLOSED_POSITIONS_SHEET_ORDER_NUMBER = 1
USD_PLN_SHEET_ORDER_NUMBER = 0
USD_PLN_HISTORY_FILE_LOCATION = './usd_pln_history.xlsx'


def select_usd_pln_sheet():
    xlsx = openpyxl.load_workbook(USD_PLN_HISTORY_FILE_LOCATION)
    sheet = xlsx.worksheets[USD_PLN_SHEET_ORDER_NUMBER]
    return sheet


usd_pln_sheet = select_usd_pln_sheet()


def select_transaction_sheet():
    xlsx = openpyxl.load_workbook(ETORO_ACCOUNT_STATEMENT_FILE_LOCATION)
    sheet = xlsx.worksheets[CLOSED_POSITIONS_SHEET_ORDER_NUMBER]
    return sheet


closed_positions_sheet = select_transaction_sheet()


def calculate_number_of_entieties(sheet):
    dimensions = sheet.dimensions
    return int(dimensions.split(':')[1].replace('R', '').replace('G', '')) - 1


numberOfTransactions = calculate_number_of_entieties(closed_positions_sheet)


def format_date(date_string):
    return date_string.split(" ")[0].replace("/", ".")


def find_usd_pln_rate_for_end_of_day(date):
    usd_pln_history_length = calculate_number_of_entieties(usd_pln_sheet)
    for i in range(usd_pln_history_length):
        cell_number = i + 2
        open_date_cell = str('A' + str(cell_number))
        if usd_pln_sheet[open_date_cell].value == date:
            closed_value_cell = str('B' + str(cell_number))
            return float(usd_pln_sheet[closed_value_cell].value)
    raise ValueError(f'Exchange rate for date not found for date: {date}')


def create_transaction_list():
    transaction_list = []
    for i in range(numberOfTransactions):
        cell_number = i + 2
        open_date_cell = str('E' + str(cell_number))
        closed_date_cell = str('F' + str(cell_number))
        open_rate = str(('J' + str(cell_number)))
        closed_rate = str(('K' + str(cell_number)))
        units_cell = str(('D' + str(cell_number)))

        open_day = format_date(closed_positions_sheet[open_date_cell].value)
        closed_day = format_date(closed_positions_sheet[closed_date_cell].value)

        open_exchange_usd_pln_rate = float(find_usd_pln_rate_for_end_of_day(open_day))
        closed_exchange_usd_pln_rate = float(find_usd_pln_rate_for_end_of_day(closed_day))

        open_rate = float(closed_positions_sheet[open_rate].value)
        closed_rate = float(closed_positions_sheet[closed_rate].value)

        units = float(closed_positions_sheet[units_cell].value)

        open_pln_value = units * open_rate * open_exchange_usd_pln_rate
        closed_pln_value = units * closed_rate * closed_exchange_usd_pln_rate

        profit = open_pln_value - closed_pln_value

        transaction_list.append([open_day,
                                 closed_day,
                                 open_rate,
                                 closed_rate,
                                 open_exchange_usd_pln_rate,
                                 closed_exchange_usd_pln_rate,
                                 profit])
    return transaction_list


def calculate_transaction_profits_in_pln(closed_transactions):
    profit = float(0)
    for i in range(numberOfTransactions):
        profit += closed_transactions[i][6]
    return profit


def main():
    closed_transactions = create_transaction_list()
    profit = calculate_transaction_profits_in_pln(closed_transactions)
    print(profit)


if __name__ == '__main__':
    main()
