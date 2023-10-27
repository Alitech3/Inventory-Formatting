#Copyright 2023, Ali Chapman, All right reserved. Message for desired use.

import csv
import sys
import os
import ctypes
ctypes.windll.kernel32.SetConsoleTitleW("Nicks For Men - Inventory Formatting")
# Categorys are to be seperated by commas
print('Intended for use with Suits, Shirts, and Shirts only')
print('*all file names are case sensitve*')
print('In order to use this program you need 2 things:')
print('>The file for adding a product to the rain software - MUST be named "addproduct.csv"')
print('>A text document with all the inventory you counted - MUST be named "inventory.txt"\n')
print('This file should follow the following format:')
print('Category ID\nProduct Name\nDefault Cost\nPrice\nSales Price\nDescription\nSizes\nXXX')
print('all of these items MUST be on their own line in the text document')
print('For the description ensure it has no new lines in it')
print('Use "XXX" when seperating product information\n')
print('When the program completes a csv file called "FormattedInventory" will be created')

response = input('type "OK" to continue\n')

if response.lower() != 'ok':
	print('restart program and be sure to type "OK" all caps')
	os.system('pause')
	quit()

InventoryType = input('Please specify the type of invetory Suits, Shoes, or Shirts:\n')

AssOp = []
CatID = None
ProTIT = None
ShoDesc = None
DC = None
P = None
SP = None
InvetoryCount = []
productinfo = []
rawData = []

#implement further check to make sure user doesnt put mismatching types and data
if InventoryType.lower() == 'suits':
	AssOp = ['34S', '34R', '36S', '36R', '38S', '38R', '38L', '40S', '40R', '40L', '42S', '42R', '42L', '44S', '44R', '44L', '46S', '46R', '46L', '48S', '48R', '48L', '50R', '50L', '52R', '52L', '54R', '54L', '56R', '56L', '58R', '58L', '60R']
elif InventoryType.lower() == 'shoes':
	AssOp = ['7', '7.5', '8', '8.5', '9', '9.5', '10', '10.5', '11', '11.5', '12', '13', '14', '15']
elif InventoryType.lower() == 'shirts':
	ShirtType = input('Please type SLIM or MODERN')
	if ShirtType.lower() == 'slim':
		NeckSize = ['13.5', '14.5', '15,5', '16.5', '17.5', '18.5']
	elif ShirtType.lower() == 'modern':
		NeckSize = ['14.5', '15.5', '16.5', '17.5', '18.5', '19.5', '20.5', '21', '22', '23', '24']
	else:
		print('You entered an invalid Fit -\n>using default fitting')
		NeckSize = ['13.5', '14.5', '15.5', '16.5', '17.5', '18.5', '19.5', '20.5', '21', '22', '23', '24']
	SleeveSize = ['32/33', '34/35', '36/37', '38/39']
else:
	print('You have entered an invalid response MUST BE EITHER SUITS, SHIRTS, OR SHOES\nrestart the program and try again')
	os.system('pause')
	quit()

try:
	with open('./addproducts.csv', newline='') as csvheaders:
		readheaders = csv.reader(csvheaders)
		for row in readheaders:
			# this takes the headers from the addproducts file and adds them to variable for use later
			headers = row
			break
except OSError as e:
	print('\n'+e.strerror)
	print('\nNo file named addproducts.csv found')
	print('Double check to make sure you downloaded the template from RainPOS website, renamed it, and put it in the same directory as this program')
	os.system('pause')
	quit()

try:
	with open('./inventory.txt', 'r') as sample:
		# this condenses the text document into one long array
		text = sample.read().split('\n')
except OSError as e:
	print('\n'+e.strerror)
	print('\nNo file named inventory.txt found')
	print('Make sure your invetory file is in the same directory as this program')
	os.system('pause')
	quit()

def mainWrite():
	writer.writerow({
		'Category IDs (Comma separate)': CatID, 
		'Product Title': ProTIT, 
		'Short Description': ShoDesc,
		'Unit of Measurement(each/per yard)': 'each',
		'Options': 'Size:' + ','.join(AssOp),
		'Assigned option values': 'Size:' + AssOp[0],
		'Default Cost': DC,
		'Price': P,
		'Sale Price': SP,
		'Inventory':  InvetoryCount[0]})
	for x in range(1, len(AssOp)):
		writer.writerow({
			'Assigned option values': 'Size:'+ AssOp[x],
			'Default Cost': DC,
			'Price': P,
			'Sale Price': SP,
			'Inventory':  InvetoryCount[x]})

def ShirtWrtie():
	writer.writerow({
		'Category IDs (Comma separate)': CatID, 
		'Product Title': ProTIT, 
		'Short Description': ShoDesc,
		'Unit of Measurement(each/per yard)': 'each',
		'Options': 'Neck Size:' + ','.join(NeckSize) + ';'+'Sleeve Size:' + ','.join(SleeveSize),
		'Assigned option values': 'Neck Size:' + NeckSize[0] + ';' + 'Sleeve Size:' + SleeveSize[0],
		'Default Cost': DC,
		'Price': P,
		'Sale Price': SP,
		'Inventory':  InvetoryCount[0]})
	for x in range(1, len(AssOp)):
		writer.writerow({
			'Assigned option values': 'Neck Size:' + NeckSize[x] + ';' + 'Sleeve Size:' + SleeveSize[x],
			'Default Cost': DC,
			'Price': P,
			'Sale Price': SP,
			'Inventory':  InvetoryCount[x]})

if InventoryType.lower() == 'shirts':
	try:
		with open('./FormattedInventory.CSV', 'w+', newline='') as invetorySheet:
			fieldnames = headers
			writer = csv.DictWriter(invetorySheet, fieldnames=fieldnames)
			writer.writeheader()
			for x in range(len(text)):
				#populates the productinfo list
				if text[x] != 'XXX':
					productinfo.append(text[x])
				if text[x] == 'XXX':
					CatID = productinfo[0]
					ProTIT = productinfo[1]
					DC = productinfo[2]
					P = productinfo[3]
					SP = productinfo[4]
					ShoDesc = productinfo[5]
					for y in range(6, len(productinfo)):
						rawData.append(productinfo[y])

	except OSError as e:
		print('\n'+e.strerror)
		print('\nAn error has occured')
		print('Make sure you do not have a FormattedInventory.csv open')
		print('or contact dev')
		os.system('pause')
		quit()

else:
	try:
		with open('./FormattedInventory.CSV', 'w+', newline='') as invetorySheet:
			fieldnames = headers
			writer = csv.DictWriter(invetorySheet, fieldnames=fieldnames)
			writer.writeheader()
			for x in range(len(text)):
				# This creates an array from the text array breaking it down by product
				if text[x] != 'XXX':
					productinfo.append(text[x])
				# this processes all the product information writes it and resets varables
				if text[x] == 'XXX':
					#print(productinfo)
					CatID = productinfo[0]
					ProTIT = productinfo[1]
					DC = productinfo[2]
					P = productinfo[3]
					SP = productinfo[4]
					ShoDesc = productinfo[5]
					#we know for a fact that everything after the 6th position in the array will be sizes so we add those to a new list
					for y in range(6, len(productinfo)):
						rawData.append(productinfo[y])
					print(rawData)
					for z in range(len(AssOp)):
						InvetoryCount.append(rawData.count(AssOp[z]))
					mainWrite()
					productinfo = []
					rawData = []
					InvetoryCount = []
	except OSError as e:
		print('\n'+e.strerror)
		print('\nAn error has occured')
		print('Make sure you do not have a FormattedInventory.csv open')
		print('or contact dev')
		os.system('pause')
		quit()

print('Task Complete. A new file named FormmattedInventory.csv has been created.')
os.system('pause')