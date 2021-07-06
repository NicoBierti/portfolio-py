# Importing Libraries
from art import *
import time
import msvcrt as m

# Functions
def cover():
    for i in range (0,55):
        print(end='-')
    print('')
    time.sleep(0.5)
    print('\n')
    tprint("Portfolio")
    time.sleep(0.5)
    print("                                            By Nico")

def intro():
    time.sleep(0.5)
    for i in range (0,55):
        print(end='-')
    print('')

    print( '--------/*WELCOME TO PORTFOLIO*\--------\n')
    time.sleep(0.5)
    print('This program is to keep track your crypto assets in a simple way. \nAs you buy you can add them in order to store information in an excel file.')
    time.sleep(0.5)
    print('\nThe portfolio app will let you:\n- Store you cryptos assets.\n- Will check for the price at the moment you add them to the porfolio and also' 
        '\n- Store the exact date you updated the information' 
        '\n- Add or substract assest as you operate in the market.\n- Delete them from the database\n- Print the portfolio in the console as you excecute the program ')
    time.sleep(1)
    print('\nWhen you add tokens please input the shot name. For example if you want to add Bitcoin, write "BTC".\n'
        '\nThese are the most used tokens:\n'
        '\nETH - Ethereum'
        '\nBNB - Binance Coin'
        '\nDOGE - Dogecoin'
        '\nLTC - Litecoin\n')
    time.sleep(1)       
    print('\nJust follow the instructions. Enjoy')
    time.sleep(1)
    
    for i in range (0,75):
        print(end='-')
    print('')

def wait():
    time.sleep(0.5)
    print("Press any key to continue")
    m.getch()
    




  
