# CUSIP_to_ISIN

Whole script in the isindb.py file. 

This script can be used to turn any list of CUSIP company identifiers into a cleaned list of the same company's ISIN identifiers.
I created my list of CUSIPs from an excel colunm, but entering a list directly would also work. The script then uses the https://www.isindb.com website to first clean the CUSIP (calculate the control digit) and then convert the cleaned CUSIP into ISIN. 

The script is in Python, and uses mainly Selenium for the web-scraping and pandas to turn the ISIN list into an excel file. 

Enjoy using this if you need to convert a long list of CUSIPs into ISINs!! I had to do this for an assignment and thaught it might save somebody the trouble of creating a similar script. 

Let me know if you find ways to make it faster.
