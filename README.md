# Google Apps: Get list of parent/child/sibling files and folders from a Google Drive Folder
A Google Apps Script for Google Sheets to list folders and files from a Google Drive folder

## What this is for:
I had a project where I needed to migrate a shared folder to a shared drive instead. I opted to make a GSheet to do all the file listings and then had the sheet generate all items in folder rather than manual entry.

(I'll try to make this into a template for quick use but for now here is a simple how to)

## How to use
1. *Open a new Google Sheet*
2. *Go to `Extensions > App Scripts`*, this will open a new tab with an empty "Untitled Project". Feel free to click that area and name your project.
3. Remove the small code they place in the project by default and *copy the contents of [code.js](code.js) from this repo and paste into the new project*. Hit save.
4. Go back to your GSheet and reload the page
5. You should now see a `List Files/Folders` menu item pop-in (takes a few seconds)
6. Click the `List File/Folders > List All Parent/Child Files and Folders` menu item
7. *Google is going to prompt you for authorization at this part, go ahead and click Authorize*
8. You'll then choose your Google Account from the list
9. You'll get a warning like so, in transparency this is my first go with app scripts so not sure how to verify, if you know please feel free to submit a pull request on this and we can update these directions. Click `Advanced > Go to [Your Projects Name] (unsafe)`.
_Note, you will porbably see your email in the blurred part of the image below._
<img width="600" alt="image" src="https://user-images.githubusercontent.com/3694594/192628104-7ce3c540-a340-4dfd-97d6-b6b9eadc675b.png">

10. You'll now see a prompt for allowing the script to access your Google Drive, *Allow* it
11. Now repeat step 6 (Click the `List File/Folders > List All Parent/Child Files and Folders` menu item)
12. Go to another tab and navigate to a folder in your GDrive and click into that folder, in the navbar you will see the ID for you folder:
`ex: https://drive.google.com/drive/u/1/folders/[folderIDHere]
13. Copy the folder Id and then paste into the prompt's input from you Google Sheet
14. Give the script a minute or two to run and you should see your things appear, you should see something like so:
<img width="1567" alt="image" src="https://user-images.githubusercontent.com/3694594/192630090-44e07963-8456-4591-897e-3fee60421b2c.png">

Thats about it! Now if you are like me, you have the option of also running the `Set Folders Formatting to Bold` menu item from the List Files/Folders menu and you should see any line item that is a folder become bold and have a light grey (#f3f3f3) background.

Other things in this file:
* Duplicate Active Tabs - Choose an active tab, choose this menu item, enter a number and viola, duplicates of the tab to whatever number you set
* Formula `=sheetnames()` creates a list of all the tabs on the workbook.
