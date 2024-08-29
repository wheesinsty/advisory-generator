# One Page Advisory Generator and Adoption Score Scraper
This is a project that automatically generates one page advisories and scrapes adoption scores from Microsoft Admin Center. It is made for my team during my internship.

The scraping module used is Playwright. The reason for using Playwright is because by the time I realised that Playwright is mainly used for testing, the project was already finished. 

The file `automate.py` generates advisories and reports the status into the excel sheet.
The file `scores.py` scrapes adoption scores into the excel sheet.

## Set-up
To use this software, ensure that 
1. Install python from the internet
   
2. In your terminal, install pip with
`python -m ensurepip --upgrade`

3. Clone this repository with
`git clone https://github.com/wheesinsty/one-page-advisory-generator.git`
                                                                                      
4. Right click the Google Chrome app and select **Properties**. Then, set the **Target:** to
`"C:\ProgramFiles\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222`.

5. Check that the repository is in the Downloads folder, then navigate to this directory in your terminal by entering (replace YOURNAME with the name of your Windows username)
`cd [path]`

6. Install the necessary dependencies with
`pip install -r requirements.txt`

7. Ensure that the excel sheet is called **test**and the advisory template is a Microsoft document called **Advisory_OnePage**                            

## Run the projects
1. Open terminal and navigate to the directory
`cd [path]`

3. Run the file you want
`python advisory.py`

## Project doesn't work?
Potential issues and how to resolve them:
1. Blank screenshots. 
    How to resolve: Your internet is too slow, try running it again. If it still doesn't work, change the time in page.wait_for_timeout()
2. Microsoft has updated the format of their Admin centre. 
    How to resolve: Adjust the code.
3. Permission error. 
    How to resolve: Ensure that the advisory template and the excel sheets are closed.

## Credits
Thank you to my team for giving me this opportunity and for helping me along the way :`)`
