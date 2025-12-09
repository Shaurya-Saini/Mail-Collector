# VIT MAILS SCRAPER
This project automates the extraction of official student email IDs from Microsoft Teams using Selenium. Given a list of student registration numbers or names, the tool logs into Teams (after manual authentication), searches for each entry, identifies the correct student from the People results, and extracts their institutional ID to generate a clean, structured email list. The script is designed to handle variations in Teams’ UI, avoid incorrect matches, and gracefully skip students who are not present on Teams.

## SetUp Instructions
1. Install dependandcies <br>
`pip install selenium pandas`

2. Setup ChromeDriver<br>
https://youtu.be/4gJAAeEPxFo?si=P3c0B0xbEcIfS6Lb - use this video to setup and install Chromedriver with selenium

3. In the `data.xlsx` file in the first column add Registration numbers of students

4. Run `main.py`

4. Using Chromedriver script will open the ms teams, log in using vit mail and click enter on the terminal running the script

5. The script will generate a new excel file including registration numbers and their mails