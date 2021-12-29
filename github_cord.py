# Copyright (c) 2020-2021 Adrian S. Lemoine
#
#   Distributed under the Boost Software License, Version 1.0. 
#   (See accompanying file LICENSE_1_0.txt or copy at 
#   http://www.boost.org/LICENSE_1_0.txt)

from github import Github
import argparse
import os
import re
import datetime
import csv

### Functions
def extract_labels(labelclass):
    str = ''
    if len(labelclass) is 0:
        return str
    else:
        for label in labelclass:
            str = str + label.name + ", "
        str = str[:-2]
        return str

# Extract the issue number from its web address
def strip_issue_number(string):
    loc = re.search(r"/issues/", string)
    return string[loc.span()[1]:]

# Calculate the number of working days in a sprint
def calc_working_days(sl):
    working_days = ""
    if sl <= 5:
        working_days = str(sl) +' days'
    elif sl % 7 is 0:
        working_days = str(sl - (sl/7)*2) +' days'
    else:
        print("Warining: Can't caluclate the number of working days. Setting working days to an empty string.")
        working_days = ""
    return working_days

# Calculate the start and end dates of a task from the due date
def calc_task_dates(ms):
    if ms == None or ms.due_on == None:
        return ('','')
    elif sprint_length == 0:
        return (ms.due_on.strftime("%m/%d/%Y"), ms.due_on.strftime("%m/%d/%Y"))
    else:
        return ((ms.due_on - datetime.timedelta(days=sprint_length-1)).strftime("%m/%d/%Y"), ms.due_on.strftime("%m/%d/%Y"))

# Remove '\u200b', '\r', or '\n'
def scrub_text(string):
    return re.sub('(\u200b)|(\r)|(\n)', '', string)

# Return milestone name
def strip_milestone(ml):
    if ml == None:
        return ''
    else:
        return ml.title

# Convert common GitHub Column names to common MS Project Names
def board_status(str):
    str=str.lower()
    if str == "to do":
        return "Not Started"
    elif str == "to-do":
        return "Not Started"
    elif str == "started (<10% complete)":
        return "Started"
    elif str == "started (<10% completed)":
        return "Started"
    elif str == "started":
        return "Started"
    elif str =="in progress":
        return "In progress"
    elif str == "review in progress":
        return "Under Review"
    elif str == "under review":
        return "Under Review"
    elif str == "reviewer approved":
        return "Under Review"
    elif str == "done":
        return "Done"
    elif str == "none":
        return "None"
    elif str == "":
        return ""
    else:
        print("String '", str, "' does not match any known board status!")
        return ""

# Assign percent complete based on board status
def percent_complete(str):
    str=str.lower()
    if str == "to do":
        return "0"
    elif str == "to-do":
        return "0"
    elif str == "started (<10% complete)":
        return "10"
    elif str == "started (<10% completed)":
        return "10"
    elif str == "started":
        return "10"
    elif str =="in progress":
        return "50"
    elif str == "review in progress":
        return "75"
    elif str == "under review":
        return "75"
    elif str == "done":
        return "100"
    elif str == "none":
        return ""
    elif str == "":
        return ""
    else:
        print("String '", str, "' does not match any known percent complete!")
        return ""

###### Program Start
### Setup
# Set up arguments
parser = argparse.ArgumentParser(description = 
    "project_logic is a program which fetches data from "
    "a GitHub Project and updates the information in a "
    "specified Excel file.")
parser.add_argument('--github_repo', help='URL to GitHub repository.')
parser.add_argument('--csv_file', nargs = '?', help = 'Path to CSV file.'
                    , default = 'project_issues_'+datetime.date.today().strftime("%Y.%m.%d")+'.csv')
parser.add_argument('--sprint_length', nargs = '?', help = 'Length of sprint in days.'
                    , default = 14)

args = parser.parse_args()

# Assign script arguments to variables
github_repo = args.github_repo
filename = args.csv_file
sprint_length = int(args.sprint_length)
access_token = os.getenv('GITHUB-TOKEN')
g = Github(access_token)

## Fetch data from GitHub
print("Fetching issues from", github_repo)
repo = g.get_repo(github_repo)
open_issues = repo.get_issues(state="all", direction="asc")
projects = repo.get_projects()

## Fetch all the cards in each project -- Takes a long time!
print("Gathering cards from each project")
issue_column = {}
open_notes = []  # Capture notes
for proj in projects:
    columns = proj.get_columns()
    for col in columns:
        proj_cards = col.get_cards()
        for card in proj_cards:
            # Capture the project and column name for each issue and note
            if card.content_url:
                issue_column[strip_issue_number(card.content_url)] = (proj.name, col.name)
            else:
                issue_column[card.id] = (proj.name, col.name)
                open_notes.append(card)

## Setup CSV File
print("Writing issues to", filename)
with open(filename, 'w', newline='') as f:
    writer = csv.writer(f)
    default_headings = ["Name","Percent_Complete","Duration","Start_Date","Finish_Date","Milestone","Board_Status","GitHub_Issue","Labels","Project"]
    writer.writerow(default_headings)
    for issue in open_issues:
        # Get the project and board status
        iproj = ''
        bs = ''
        if str(issue.number) in issue_column:
            iproj = issue_column[str(issue.number)][0]
            bs = issue_column[str(issue.number)][1]
        elif issue.state == 'closed':
            bs = 'Done'
        task_dates = calc_task_dates(issue.milestone)
        csvrow = [scrub_text(issue.title)
                  , percent_complete(bs)
                  , calc_working_days(sprint_length)
                  , task_dates[0], task_dates[1]
                  , strip_milestone(issue.milestone)
                  , board_status(bs), issue.number
                  , extract_labels(issue.labels)
                  , iproj]
        writer.writerow(csvrow)
    for note in open_notes:
        # Get the project and board status
        iproj = issue_column[note.id][0]
        bs = issue_column[note.id][1]
        csvrow = [scrub_text(note.note)
                  , percent_complete(bs)
                  , calc_working_days(sprint_length)
                  , '', ''
                  , ''
                  , board_status(bs), note.id
                  , ''
                  , iproj]
        writer.writerow(csvrow)
print("Complete!")
