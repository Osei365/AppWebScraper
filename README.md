# Appscraper
### This scraper was built to scrape  the [App Exchange](https://appexchange.salesforce.com) website. 
### The major tools used in scraping the website was the Selenium Webdriver. It was used to automate every process that led to the successful scraping of the website. Other tools, packages or frameworks that were used are listed below:
- Xlsxwriter: To create an Excel workbook to save the data being scraped
- Openpyxl: This was used to load each row into the data iteratively
- time: This is a python package that is used in the manipulation of time based data structures. One of it's method *sleep* was used to make the program rest a little before scraping further at strategic intervals

### The website is structured in such a way that the apps have been grouped into categories and then each category has more than 100 apps. 
### Therefore, the first challenge in scraping the website was to click each category, get all the apps present in the page of each category and return to the main page to continue with another category. 

### The code above did justice to that. The xpath for each category was identified, gathered and assigned to a python iterable. Afterwards a for loop was used to iterate through each category till all the apps for each category was scraped.
### The categories were: 
- Finances, 
- analytics, 
- Human Resources,
-  Sales,
-   Enterprise
- Customer Service
- IT and Admin
- Marketing
- Integration
- Salesforce Labs
### They were ten(10) top categories in all.

### Another major milestone was to completely load the page of each category. The apps weren't fully displayed on the screen. A _show more_ button was used to load more apps continuously for a period of time. The program accounted for this as well.

### After the apps were fully loaded on each category page, each app had to be clicked to retrieve the fields that was needed. They were fifteen (15) that were required for each app. These fields formed the column names on the Excel Workbook.

### This piece summarizes the technicalities in scraping this particular website (https://appexchange.salesforce.com). please, feel free to suggest ways I can upgrade the quality of the code.

### Thank you for reading
