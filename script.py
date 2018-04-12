from O365 import Inbox
from datetime import datetime, timedelta
import argparse, json, re

# define args
parser = argparse.ArgumentParser()
parser.add_argument("dir", help="input directory for email download")
args = parser.parse_args()
path = args.dir

# connect to inbox
auth = ("qa.ex@office365.ecknhhk.xyz", "ew68I7W52p*W")
inbox = Inbox(auth, getNow=False)
# define 24-hour filter
one_day_back = (datetime.utcnow() - timedelta(hours=24)).strftime("%Y-%m-%dT%H:%M:%SZ")
inbox.setFilter("DateTimeReceived ge " + one_day_back)
inbox.setOrderBy("DateTimeReceived desc")
# download emails (using 1000 because defaults to only 10)
inbox.getMessages(1000)
# if nothing within last 24 hours, get recent 10 emails
if len(inbox.messages) is 0 and len(inbox.errors) is 0:
    inbox.setFilter("")
    inbox.getMessages(10)

# save each email locally (filename = [index]_[DateTimeReceived])
index = 0
for email in inbox.messages:
    received = email.json['DateTimeReceived'].translate({ord(k):None for k in u'-:'})
    filename = str(index) + '_' + received
    index+=1
    with open(path + '/' + filename + '.txt', 'w') as outfile:
        json.dump(email.json, outfile)

print("Done. Downloaded " + str(index) + " emails.")
