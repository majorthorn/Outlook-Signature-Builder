Outlook Signature Builder

This powershell script started off as an internal project that I was allowed to work on inbetween helpdesk tickets where i work. However, due to time constraints and time required to implement a working script the project was stopped. I got permission to scrub the company name off of the script and provide it and open source script. 

The reason why this script exists is because there is a VBScript that has made the rounds on the internet that we used for generating outlook signatures for our users. This was all in an attempt to standardize our email signatures to make sure that we had the required contact information in all user signatures. The VBScript that was used is rather old and not user friendly when it comes to doing manual builds of signatures. So I wanted to make a powershell script that did basically all of the same functions but did not interface directly with Outlook. We didnt want to interface directly with Outlook because the signatures would need to be added to the user's Outlook manually. To do this we just needed a Word document that they could copy and paste into their signature form. The functions of this script can be performed with a mailmerge in Word itself but I found that if i needed to generate multiple word documents the time required was higher than if I was able to automate the task.


This is completely a work in progress and is mostly for reference at this point due to the requirements of the script.


# Requirements
* Admin access to an Active Directory Domain
* Windows 10
* Microsoft Office (Word, Outlook)

This Repo uses GitFlow from GitKraken.
