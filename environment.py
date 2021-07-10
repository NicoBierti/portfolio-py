# Importing Libraries
from openpyxl.styles import Font
import cryptocompare
from art import *
import msvcrt as m
import os
from openpyxl import load_workbook

# Set working directory and workbook
current_working_directory = os.getcwd()
os.chdir(current_working_directory)
workbook = load_workbook('Portfolio.xlsx')
workbook.active.title = " Cryptofolio "

# Banner in console, Intro and description of the program
def intro():
    for i in range (0,55):
        print(end='-')
    print('')
    print('\n')
    tprint("Portfolio")
    print("                                            By Nico\n")
    for i in range (0,55):
        print(end='-')
    print('')
    print( '--------/*WELCOME TO PORTFOLIO*\--------\n')
    print('This program is to keep track your crypto assets in a simple way. \nAs you buy you can add them in order to store information in an excel file.')
    print('\nThe portfolio app will let you:\n- Store you cryptos assets.\n- Will check for the price at the moment you add them to the porfolio and also' 
        '\n- Store the exact date you updated the information' 
        '\n- Add or substract assest as you operate in the market.\n- Delete them from the database\n- Print the portfolio in the console as you excecute the program ')
    print('\nWhen you add tokens please input the shot name. For example if you want to add Bitcoin, write "BTC".\n'
        '\nThese are the most used tokens:\n'
        '\nETH - Ethereum'
        '\nBNB - Binance Coin'
        '\nDOGE - Dogecoin'
        '\nLTC - Litecoin\n')    
    print('\nJust follow the instructions. Enjoy')
    for i in range (0,75):
        print(end='-')
    print('')
    print("Press any key to continue")
    m.getch()

# Get Cryptocurrencies prices in USD using Cryptocompare.com API.
def api_call():
    cryptocompare.cryptocompare._set_api_key_parameter('59711b0f290e155ff93aa6e10f183764ea6e167caab9379aadc88928308155c1')

# Column Headers
def formatting():
    workbook.active['A1'] = 'Asset'
    workbook.active['B1'] = 'Price (USD)'  
    workbook.active['C1'] = 'Amount'
    workbook.active['D1'] = 'Balance (USD)'
    workbook.active['E1'] = 'Date'
    workbook.active['F1'] = 'Time'

    bold = Font( bold = True)
    for cell in workbook.active["1:1"]:
        cell.font = bold        



