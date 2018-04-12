from O365 import Inbox
from datetime import datetime, timedelta
import argparse, json, re
USERNAME = "qa.ex@office365.ecknhhk.xyz"
PASSWORD = "ew68I7W52p*W"

# main flow
def main(args):
    # connect to inbox and download messages
    inbox = login_and_download(USERNAME, PASSWORD)
    # upload to s3 or download locally
    download_locally(inbox, args.dir)


# handle args
def parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("dir", help="input directory for email download")
    args = parser.parse_args()
    return args


# connect to inbox and download messages
def login_and_download(username, password):
    auth = (username, password)
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
    return inbox

# save each email locally (filename = [index]_[DateTimeReceived])
def download_locally(inbox, path):
    index = 0
    for email in inbox.messages:
        received = email.json['DateTimeReceived'].translate({ord(k):None for k in u'-:'})
        filename = str(index) + '_' + received
        index+=1
        with open(path + '/' + filename + '.txt', 'w') as outfile:
            json.dump(email.json, outfile)

    print("Done. Downloaded " + str(index) + " emails.")


if __name__ == "__main__":
    args = parse_arguments()
    main(args)
