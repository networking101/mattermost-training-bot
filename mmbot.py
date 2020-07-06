import requests
import json
import copy
import openpyxl
import datetime
import difflib
import sys
import os


######################
### Configurations ###
######################

# Ratio when comparing usernames from mattermost and training report
# When the ratio is closer to 0, the user will be asked to deconflict names that differ more
# When the ratio is closer to 1, more names will be skipped and not double checked by user
# If the program is skipping users, lower the ratio.  If the program is asking too often, increase the ratio
userRatio = .75
# Ratio used when comparing training names from configs dictionary and training report
# If you keep getting asked for this, consider changing the names in the "trainings" dictionary to match the training report
trainingRatio = .66

# This dictionary contains the training are checked for concurrency.  The key should match the headers in the excel doc.
# The value is the number of years before expiration
trainings = {
    "Cyber Awareness": 1,
    "Free Ex Religion": 3,
    "Force Protection": 1,
    "Combat Trafficking Persons": 3,
    "Unauth Disclos": 1,
    "Intel Oversight": 1
}

# Tokens and IDs needed to interact with the API
configs = {
    # Used to authenticate the bot
    "bot_access_token": "w8k5cejwnp8ifk66zbexzxprmc",
    # Bot identifier.  Used to set up direct message channel between bot and user
    "bot_token_id": "994c6te3j3dszkqiiy8t6gp1ze",
    # Channel identifier.  Used to pull user list.  Make sure that this channel has all active members that will need to be notified of trainings
    "main_channel": "5bzj3zsz6brmjgat8gx3ejudge",
    "renew_time": trainings
}
### End Configurations ###

########################
### Global Variables ###
########################

currTime = datetime.datetime.now().timestamp()
# number of seconds in a year
year = 31536000
# number of seconds in 31 days
month = 2678400

# Stores the headers from training report.  Key is the column location and value is the training name
wsHeaders = {}
# Stores user information from mattermost
users = []
# Stores user information from training report
pwnieList = []

# Stores the location of Names in training report
nameColumn = -1

# set authorization token
headers = {
    "Authorization": "Bearer " + configs["bot_access_token"]
}


### End Global Variables ###


# Function used to check similarity between strings.  Some users names from mattermost will not match the training report.
# If there is a difference, the similarity will be calculated as a ratio.  An exact match will always pass.  Anything below
# the ratio limit will fail. Anything between the ratio limit and 1 will be verified with the script user.
def checkString(origStr, newStr, simRatio):
    similarity = difflib.SequenceMatcher(None, origStr.lower(), newStr.lower()).ratio()

    if similarity == 1:
        return True
    if similarity >= simRatio:
        print("\n" + str(round(similarity * 100, 2)) + "% match")
        while True:
            retBool = input("Is " + newStr + " the same as " + origStr + "? [y or n]: ")
            print("")
            if retBool.lower() == "y":
                return True
            if retBool.lower() == "n":
                return False
            print("Does not compute. Try again")

    return False


# Get an updated list of active users from mattermost
def getPwnies():
    global headers
    global users
    
    
    # Required for API request
    params = {
        "page": "0",
        "per_page": "60",
        "in_channel": configs["main_channel"]
    }

    # grab first page of users
    r = requests.get("https://cyberpwnies.com/api/v4/users", headers=headers, params=params)
    users = r.json()

    # if there are more pages, keep requesting until all users are pulled
    while len(r.json()) == 60:
        params["page"] = str(int(params["page"]) + 1)
        r = requests.get("https://cyberpwnies.com/api/v4/users", headers=headers, params=params)
        users += r.json()

    print("Found " + str(len(users)) + " pwnies on mattermost")


# Use openpyxl to get data from the training report
def openTrainingDoc(docName):
    global wsHeaders
    global pwnieList

    try:  #TODO
        wb = openpyxl.load_workbook(docName, data_only=True)
    except:
        print("Error opening " + docName)
        exit()

    # Search for tab named "Squadron All".  If this changes often between reports, I can change this to always pull the first tab
    try:
        ws1 = wb["Squadron All"]
    except:
        print("Could not find worksheet labled \"Squadron All\"")
        exit()

    # parse through column headers
    #    1) Make sure the trainings match with config.json
    #    2) Find the "Flight" column
    #    3) Find the "Name" column
    fltColumn = -1
    global nameColumn

    # cycle through cels in top row
    for cell in ws1[1]:
        if cell.value.lower() == "flight":
            fltColumn = cell.column - 1
            continue
        if cell.value.lower() == "name":
            nameColumn = cell.column - 1
            continue
        for tr in configs["renew_time"]:
            # check if training name from training report matches name from "configs".  *NOTE* We must store the exact name from the training report in wsHeaders for gatherUserReport function
            if checkString(cell.value, tr, trainingRatio):
                wsHeaders[cell.column] = tr
            #else:
            #    print("If this program is not finding the correct names, consider updating the training names in \"configs\" to match the training report")
    if fltColumn == -1 or nameColumn == -1:
        print("Could not find \"Flight\" column or \"Name\" column in \"Squadron All\" worksheet")
        exit()

    # get all pwnies who are overdue or upcomming from training report
    count = 0
    for row in ws1.iter_rows():
        # need first part to filter out records with empty columns
        if row[fltColumn].value and (row[fltColumn].value.lower() == 'c' or row[fltColumn].value.lower() == 'z' or row[fltColumn].value.lower() == 'ado'):
            count += 1
            pwnieList.append(row)
    print("Found " + str(count) + " Z and C flight members listed in training report")


# Sends the training information in a direct message
def sendDM(userID, memberStats):
    global headers

    # Get the direct message channel between the user and the bot
    payload = [configs["bot_token_id"], userID]
    r = requests.post("https://cyberpwnies.com/api/v4/channels/direct", headers=headers, json=payload)

    dmID = r.json()['id']

    # Generate the message to the user.  Depends on training that is overdue or upcomming
    message = "This is an automated message\n"
    if memberStats["overdue"]:
        message += "\n\nYou are overdue on the following trainings:\n\n"
        for x in memberStats["overdue"]:
            message += "   - " + x + "\n"
    
    if memberStats["upcomming"]:
        message += "\n\nYou are about to be overdue on the following trainings:\n\n"
        for x in memberStats["upcomming"]:
            message += "   - " + x + "\n"
    
    message += "\nIf you believe this is incorrect, reach out to your flight training manager"

    # Send message to user
    payload = {"channel_id": dmID, "message": message}
    r.status_code = 0
    r = requests.post("https://cyberpwnies.com/api/v4/posts", headers=headers, json=payload)
    print(memberStats["name"] + "    HTTP Response Code : " + str(r.status_code))


# Find which record from training report matches which user from mattermost
def checkRecordMatch(memberStats):
    global users

    possibleUsers = []

    # First, loop through each user from mattermost and identify if a user has an exact match
    for each in users:
        if memberStats["name"].lower() == (each["last_name"] + ", " + each["first_name"]).lower():
            # Send direct message to user
            sendDM(each["id"], memberStats)
            
            # Remove this user from records to cut down on asks
            users.remove(each)

            # If users match, return true to remove it from newPwnieList
            return True

        possibleUsers.append(each)

    # If a perfect match is not found loop through possibleUsers to check
    for each in possibleUsers:
        if checkString(memberStats["name"], each["last_name"] + ", " + each["first_name"], userRatio):
            # Send direct message to user
            sendDM(each["id"], memberStats)

            # Remove this user from records to cut down on asks
            users.remove(each)

            # If users match, return true to remove it from newPwnieList
            return True
    
    # If no match found, return false
    return False


# Report on users who were found in training report but were not sent a message
def aar(count, newPwnieList):
    print("\n\nComplete!\n\n" + str(count) + " messages sent")

    if newPwnieList:
        print("The following users were found in the training report but did not receive messages\n")
        for user in newPwnieList:
            print(user)
        print("\nThis means that they are either not in the mattermost announcements channel or their name in mattermost was too different from their name in the training report.")
        print("Check to make sure their names are correct or consider adjusting the \"userRatio\" in \"configs\".")
    else:
        print("All messages sent!")


# For each pwnie in the training report, pull the trainings that are overdue and upcomming
def gatherAndSendUserReport():
    global nameColumn
    global pwnieList
    newPwnieList = []
    count = 0

    # Cycle through pwnies in training report
    for row in pwnieList:
        # Data structure sent to sendDM function
        memberStats = {"name": row[nameColumn].value, "overdue": [], "upcomming": []}
        for cell in row:
            # Check if cell value is a date.  If so, we can assume this is a "last completed date" for a training
            if isinstance(cell.value, datetime.datetime):
                # Check if overdue.  Current epoch - (seconds in a year * years to expire).  If this value is greater than the training report date, training is overdue.
                if currTime - (configs["renew_time"][wsHeaders[cell.column]] * year) > cell.value.timestamp():
                    memberStats["overdue"].append(wsHeaders[cell.column])
                # Check if upcomming.  Current epoch - ((seconds in a year * years to expire) - 1 month).  If this value is greater than the training report date, training is within 1 month of going overdue.
                elif currTime - (configs["renew_time"][wsHeaders[cell.column]] * year) + month > cell.value.timestamp():
                    memberStats["upcomming"].append(wsHeaders[cell.column])
        # Check that this user is in mattermost and send this user their report in a direct message
        if not checkRecordMatch(memberStats):
            newPwnieList.append(row[nameColumn].value)
        else:
            count += 1

    # Print report on final status
    aar(count, newPwnieList)


# For debugging purposes. Send to a single user.
def individualSend(username):
    global nameColumn
    found = False

    # Cycle through pwnies in training report
    for record in pwnieList:
        if checkString(record[nameColumn].value, username, userRatio):
            found = True
            # Data structure sent to sendDM function
            memberStats = {"name": record[nameColumn].value, "overdue": [], "upcomming": []}
            for cell in record:
                # Check if cell value is a date.  If so, we can assume this is a "last completed date" for a training
                if isinstance(cell.value, datetime.datetime):
                    # Check if overdue.  Current epoch - (seconds in a year * years to expire).  If this value is greater than the training report date, training is overdue.
                    if currTime - (configs["renew_time"][wsHeaders[cell.column]] * year) > cell.value.timestamp():
                        memberStats["overdue"].append(wsHeaders[cell.column])
                    # Check if upcomming.  Current epoch - ((seconds in a year * years to expire) - 1 month).  If this value is greater than the training report date, training is within 1 month of going overdue.
                    elif currTime - (configs["renew_time"][wsHeaders[cell.column]] * year) + month > cell.value.timestamp():
                        memberStats["upcomming"].append(wsHeaders[cell.column])
            # Check that this user is in mattermost and send this user their report in a direct message
            checkRecordMatch(memberStats)

    if not found:
        print("Could not find user in training report.  No data to send")


def printHelp(arg0):
    print(
        "Mattermost Training Notification Bot\n\n"
        "Usage: python3 {name} TrainingReport.xlsx\n\n"
        "Requirements:\n"
        " - Training report must have \"Squadron All\" worksheet (tab)\n"
        " - Worksheet must have \"Name\" and \"Flight\" columns\n"
        " - Worksheet names must be in the \"Last, First\" format\n\n"
        "Configuration changes can be made.  See first few lines of program\n".format(name=arg0))

def main(argv):

    if len(argv) < 2:
        print("Missing argument")
        return
    if argv[1] == "-h" or argv[1] == "--help":
        printHelp(argv[0])
        return
    if not os.path.exists(argv[1]) or not (argv[1][-5:] == ".xlsx" or argv[1][-5:] == ".xlsm"):
        print("Incorrect file. Must be .xlsx or xlsm in path")
        printHelp(argv[0])
        return

    # Get dictionary of members from mattermost.  Stored in "users" global variable
    getPwnies()

    # Get list of users in C and Z flight from training report.  Stored in pwnieList
    # Get dictionary if trainings from training report.  Stored in wsHeaders
    openTrainingDoc(argv[1])

    ready = input("\nReady to start sending reports? [y or n]").lower()
    while ready != 'y':
        if ready == 'n':
            print("Abort")
            return
        ready = input("Does not compute. Try again [y or n]").lower()

    # Get the user training report for each user and send a direct message to each.  User is removed from pwnieList if message sent
    gatherAndSendUserReport()
    
    # Used for debugging.  Send a message to a specified user.  Format must me "Last, First"
    #individualSend("kowpak, michael")


if __name__ == "__main__":
    main(sys.argv)
