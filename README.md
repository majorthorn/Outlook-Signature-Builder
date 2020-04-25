# Outlook Signature Builder

This script was originally started as an attempt to port a VBScript that has made the rounds on the internet as a way to created a standardized outlook signature. The VBScript that was used is rather old and not user friendly when it comes to doing manual builds of signatures. I wanted to make a powershell script that did basically all of the same functions but did not interface directly with Outlook. To do this we just needed a Word document that they could copy and paste into their signature form. The functions of this script can be performed with a mailmerge in Word itself but I found that if i needed to generate multiple word documents the time required was higher than if I was able to automate the task.

This script is most reference at this point due to the requirements of the script. Being that it requires Admin access to Active Directory, Windows 10, and an licensed version of Microsoft office specifically Word and Outlook. Also, I use Linux in my personal life making Powershell Scripting not as easy.


### Requirements
* Admin access to an Active Directory Domain
* Windows 10
* Microsoft Office (Word, Outlook)

This Repo uses GitFlow from GitKraken.
