#! python3
from openpyxl import Workbook, load_workbook
import pyperclip
import datetime, smtplib, sys
from email_functions import send_email_with_attachment
import pandas as pd
import numpy as np
import webbrowser

FileName = 'OurRecipesBook.xlsx'
List_Filename = f'ShoppingList{str(datetime.date.today())}.txt'

##wb = load_workbook(FileName)

# Send to list:
##SendToList = ['avouzas@gmail.com','elsatmv@gmail.com']
SendToList = ['avouzas@gmail.com']

df_recipes = pd.read_excel(FileName, sheet_name='Recipes')
df_recipes_selected = df_recipes.dropna()

# take the recipes to fetch each recipe
list_recipes = list(df_recipes_selected.Recipes.values)
df_recipes_all = pd.DataFrame()
found_in = dict()
for recipe in list_recipes:
    df_tmp = pd.read_excel(FileName, sheet_name=recipe)
##    if recipe == 'Week smoothies':input('adfadfsa')
    # keep the information of where every ingredient is found
    for ing in list(df_tmp.Ingredients.values):
        if not found_in.get(ing,False):
            found_in.setdefault(ing, [recipe])
        else:
            found_in[ing].append(recipe)
    
    # multiply with the number of recipes
    df_tmp.Quantity = df_tmp.Quantity * df_recipes_selected[df_recipes_selected.Recipes == recipe].UserSelection_Times.values[0]
    # append all recipes into a common dataframe
    df_recipes_all = df_recipes_all.append(df_tmp)

# fix the df that controls the found in recipe
for k,v in found_in.items():
    found_in[k]='/'.join(v)
df_t = pd.DataFrame.from_dict(found_in, columns=['Found_in'], orient='index')
df_t.reset_index(inplace=True)
##input('STOP!!!')

# add all ingredients together to form the final list
df_piv = pd.pivot_table(df_recipes_all, values='Quantity', index='Ingredients', aggfunc=np.sum)
##df_recipes_all.Ingredients.unique()
df_piv_2 = pd.merge(df_piv, df_recipes_all.drop_duplicates(subset=['Ingredients']), how='left', on='Ingredients', suffixes=('','_y'))

df_lookup = pd.read_excel(FileName, sheet_name='LookUpLists')
df_piv_3 = pd.merge(df_piv_2, df_lookup, how='left', left_on='Sector', right_on='Categories')
df_piv_3.sort_values(by='Order', inplace=True)
df_piv_4 = pd.merge(df_piv_3, df_t, how='left', left_on='Ingredients', right_on='index')

df_piv_4.to_excel(f'{str(datetime.date.today())}_Mains_Shopping_List.xlsx', index=False,columns=['Ingredients','Quantity','Unit','Categories','Order','Found_in'])

df_piv_4.to_clipboard(index=False, header=False, columns=['Ingredients','Categories','Order','Quantity', 'Unit','Found_in'])
print('The final table is copied in the clipboard. Please merge with Google Sheet list')
supermarketList = 'https://docs.google.com/spreadsheets/d/1SGHV_ETmU6OjC7g14Pk4J4Il-tq-pTXl-IEZVqxHE20/edit#gid=1386834576'
##webbrowser.open(supermarketList, new=1)
webbrowser.open(supermarketList, new=1)












