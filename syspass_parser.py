#libraries
from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import clipboard
import openpyxl

#opening file
workbook = openpyxl.Workbook()
workbook.save('pswds.xlsx')
sheet = workbook['Sheet']
sheet['A1'] = 'name'
sheet['B1'] = 'category'
sheet['C1'] = 'user'
sheet['D1'] = 'url'
sheet['E1'] = 'password'

#reading string vars
print('Greetings!')
username_string = input('Enter username:')
password_string = input('Enter password:')
company_name = input('Enter company name:')

#downloading chrome webdriver and going to our site
browser = webdriver.Chrome(ChromeDriverManager().install())
browser.get('https://newclue.shortcut.ru/')

#preparing actions
action = ActionChains(browser)

#find elements
username = browser.find_element_by_name('user')
password = browser.find_element_by_name('pass')
login_button = browser.find_element_by_id('btnLogin')


#enter credentials
username.send_keys(username_string)
password.send_keys(password_string)
login_button.click()

#waiting for page to render
sleep(3)

#clicking on "Select Client" and entering company name
client_selector = browser.find_elements_by_xpath("//*[@class='selectize-input items not-full has-options']")[0]
action.move_to_element(client_selector).move_by_offset(4,0).click().perform()
action.send_keys(company_name).send_keys(u'\ue007').perform()

#doing that for every page
while True:
	#waiting for page to render
	sleep(1)

	#preparing "copy password" buttons
	copy_password_buttons = browser.find_elements_by_xpath("//*[@class='btn-action material-icons btn-action clip-pass-button mdl-color-text--indigo-A200']")

	#preparing BeautifulSoup
	requiredHtml = browser.page_source
	soup = BeautifulSoup(requiredHtml, 'html5lib')
	cards = soup.find_all('div', {'class': 'account-label round shadow'})

	#counter used for cells and copy_pasword_buttons
	i = 0
	
	#for every card copy data to cells
	for card in cards:
		
		name = card.find('div', {'class': 'field-account field-text label-field'}).find('div', {'class': 'field-text'})
		if name != None: 
			name = name.text.strip()
		else:
			name = ""
		
		category = card.find('div', {'class': 'field-category field-text label-field'}).find('div', {'class': 'field-text'})
		if category != None: 
			category = category.text.strip()
		else:
			category=""
		
		user = card.find('div', {'class': 'field-user field-text label-field'}).find('div', {'class': 'field-text'})
		if user != None: 
			user = user.text.strip()
		else:
			user=""
		
		url = card.find('div', {'class': 'field-url field-text label-field'}).find('div', {'class': 'field-text'})
		if url != None: 
			url = url.text.strip()
		else:
			url=""

		browser.execute_script("arguments[0].scrollIntoView()", copy_password_buttons[i])
		copy_password_buttons[i].click()

		sheet['A'+str(i+2)] = name
		sheet['B'+str(i+2)] = category
		sheet['C'+str(i+2)] = user
		sheet['D'+str(i+2)] = url
		sheet['E'+str(i+2)] = clipboard.paste()

		clipboard.copy("")
		
		i+=1
	
	#check if there are more pages
	try:
		next_page = browser.find_element_by_id('btn-pager-next')
		next_page.click()

	except Exception:
		break

workbook.save('pswds.xlsx')
browser.quit()

print('Passwords saved in pswsds.xlsx at script\'s folder')
print('Bye!')