# One Page Advisory Generator
This is a project that automatically generates one page advisories and scrapes adoption scores from Microsoft Admin Center. It is made for my team during my internship.

The scraping module used is Playwright. The reason for using Playwright is because by the time I realised that Playwright is mainly used for testing, the project was already finished. 

The file `automate.py` generates advisories and reports the status into an excel sheet.

## Set-up
To use this software, ensure that 
1. Install `python` and `git` from the internet
   
2. In your terminal, install pip with
   ```
   python -m ensurepip --upgrade

4. Clone this repository with
   ```bash
   git clone https://github.com/wheesinsty/one-page-advisory-generator.git
                                                                                      
5. Right click the Google Chrome app and select **Properties**. Then, set the **Target:** to
   ```bash
   C:\ProgramFiles\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222

6. Check that the repository is in the Downloads folder, then navigate to this directory in your terminal by entering 
   ```bash
   cd [path]

8. Install the necessary dependencies with
   ```bash
   pip install -r requirements.txt
                     

## Run the projects
1. Open terminal and navigate to the directory
   ```bash
   cd [path]

2. Run the file 
   ```bash
   python advisory.py

4. Enter the path to the excel sheet

5. Enter the path for the generated reports

6. If you want to start generating reports from scratch, enter "yes"

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
