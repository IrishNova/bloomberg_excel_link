from openpyxl import Workbook
import string


def generate_alternate_alphabets_upper():

    all_letters = string.ascii_lowercase
    every_other_letter = all_letters[::2]
    every_other_letter_list = [letter.upper() for letter in every_other_letter]

    for letter in every_other_letter:
        double_letter = (letter * 2).upper()
        every_other_letter_list.append(double_letter)

    return every_other_letter_list


def user_entry():
    save_name = input('Enter XLSX name: ')
    tickers = input('Enter List of Tickers: ')
    tickers = tickers.split(',')
    return save_name, tickers


def builder(sn, tk):

    wb = Workbook()
    ws = wb.active

    letter = generate_alternate_alphabets_upper()
    for i in range(len(tk)):
        col = f"{letter[i]}1"
        ws[col] = f'=BDH("{tk[i]} US EQUITY", "NET_INCOME", "01/01/1950", "11/23/2023" )'

    wb.save(f"/path/{sn}.xlsx")
    print(f"\n{sn} created\n--------------------")


if __name__ == '__main__':
    while True:
        one, two = user_entry()
        builder(one, two)

# test
# Airlines
# ALK, ALGT, AAL, DAL, HA, JBLU, LUV, SAVE, UAL, SKYW, MESA
