import pandas as pd
import pyautogui as pya, time
# excelfile = input("What is your the file's name (case sensitive)?: ")
df = pd.read_excel('/Users/jorgealfaro/Downloads/yimappings2.xlsx', 1, usecols=[0]) #second sheet, first column

dflist = df['COLUMN'].tolist() # puts everything under "column" column in a list
print(len(dflist))


time.sleep(5) # allows time to navigate to RAD tool & place cursor on conditional 

for CODE in dflist:
	pya.click(pya.position()) #function runs through list and creates RAD conditionals

pya.press('tab') 
pya.press('tab') #navigates down to the first conditional
pya.press('tab')



for CODE in dflist: # loops through dflist and populates conditionals according to code
	
	pya.typewrite( "{{SubsidiaryPlanId}} = " )
	pya.press('tab')
	pya.typewrite('"' + CODE + '"')

	if CODE != dflist[len(dflist)-1]: #if it's not the last code, navigate down to the next conditional
		for x in range(3): pya.press('tab')
		
