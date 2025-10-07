# How to Script Tasks using Python

By: Spencer Ord

10/2/25

## Introduction

Maybe you've seen "scripting in powershell or python preffered" line in a job posting, or maybe you've heard of it but don't know what it is. Scripting will make your life so much easier when you learn how to apply it to the field you're in. It is the process of taking an often repeated task and turning it into a script, or a program, that will either run itself or that you can run in a few commands. It saves time and energy that you can put towards something else. Unlike normal coding, which usually involves building an application from the ground up, scripting performs a specific task using other applications. 

To get a better understanding of what scripting can do, I will give you some examples:
- automatically backing up files every day on your operating system
- extract a set up data from a server daily, weekly, yearly, etc.
- sending out weekly reminder emails or updates
- automatically download files from the web

These examples are just the tip of the iceburg. There are so many possibilities! Tired of writing a weekly update email to your boss? You can make a script for that! Tired of downloading certain files every day when you get to work? You can write a script for that too! If you're in the cybersecurity field, you can write a script that scans logs for failed login attemptes. In IT operations, you can have a script rename 1000 files all at once! If you do data analysis, you can automatically take data from a csv file and create a chart. The possiblilities are really endless.

Python is a popular language for scripting for several reasons. First, the python language is one of the easier languages to read, making for quicker learning and simpler work. Second, python has a bunch of pre-built libraries that you can use, whether it be for automation, networking or data. Lastly, python works on any platform, whether it be Windows, Linux, or Mac. 

For our purposes today, I will demonstrate how to parse through a CSV file and present it as a chart, graph, or report in excel. Before getting there though, we need to follow a pattern to ensure we are prepared. You can follow this pattern to identitfy what tasks you could write a script for.

1. Identify a repetitive task that you do a lot
2. Research and pick a library/tool that python has built in
3. Write your script
4. Test/debug it in a safe environment
5. Schedule your script

It isn't too an overly complicated process! Making sure you are prepared ahead of time wille ensure writing your script is fast and efficient. Testing it in a safe environment ensures you don't mess anything up in your work or home environment. Scheduling is the last part. It ensures you don't have to worry about it again. Once we are done creating our excel report script, you won't have to worry about it every again. It will automatically create it daily, weekly, or monthly for you with the correct numbers.

Scripting doesn't have to be overly complicated! So lets dive in.

## The Walkthrough

Today I want to familiarize you with several of the capabilites that python has. To do this, this walkthorugh will focus on how to write a script that will do the following:

- take data from an excel file
- format it nicely
- send the nicely formatted file in an excel file to someone
- automate it to repeat weekly

This will familiarize you with how to use python, several of the libraries, and the process of scheduling it to run automatically. 

### Initial Setup

This lab will give the main instructions for Windows OS users, however there will be notes when instructions differ for macOS or linux.

If you don't already have VSCode installed, download it. here is a link: [VSCode](https://code.visualstudio.com/download)

**Create a new folder** to contain all your files wherever you desire, and name it "Walkthrough" or something else if you desire. 

Now open VSCode, and click on the "Open Folder" button. Navigate to your newly made folder, and click select folder.

### Create Files
In the lefthand pannel that appears, you should see the word "WALKTHROUGH" (or the name of your folder you created). Found next to the name, click on the icon that look like a piece of paper with a + in the bottom right corner. This creates a new file. Name it **.env** and click enter.

Now click on the same icon, and this time name the file **weekly_email.py**.

If your side pannel matches the following picture, you've done it right so far.


