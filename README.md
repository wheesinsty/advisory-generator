# One Page Advisory Generator 
This is a project that automatically generates the one page advisories is made for my team during my internship.

The scraping module used is Playwright. The reason for using Playwright is because by the time I realised that Playwright is mainly used for testing, the project was already finished. 

## Set-up
To use this software, ensure that 
1. Install python from the internet
2. In your terminal, install pip with
`python -m ensurepip --upgrade`
3. Clone this repository with
`git clone https://github.com/wheesinsty/advisory-generator.git`                                                                                             
4. Right click the Google Chrome app and select **Properties**. Then, set the **Target:** to
`"C:\ProgramFiles\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222`. 
5. Check that the repository is in the Downloads folder, then navigate to this directory in your terminal by entering (replace YOURNAME with the name of your Windows username)
`C:/Users/YOURNAME(SUPERHU/Downloads/AdvisoryAutomate/`
6. Install the necessary dependencies with
`pip install -r requirements.txt` 
7. Ensure that the excel sheet is called **test**and the advisory template is a Microsoft document called **Advisory_OnePage**                            

## Install the project
1. Ensure that you are in the AdvisoryAutomate directory (replace YOURNAME with the name of your Windows username)
`C:/Users/YOURNAME(SUPERHU/Downloads/AdvisoryAutomate/`
2. Install the project
`python automate.py install`
3. Run the project
`python automate.py`

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
