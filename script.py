from O365 import Inbox
from datetime import datetime, timedelta
import argparse, json, re, requests, os, boto3
USERNAME = "qa.ex@office365.ecknhhk.xyz"
PASSWORD = "ew68I7W52p*W"
BUCKET_NAME = "interview-exercises"

# main flow
def main(args):
    # connect to inbox and download messages
    inbox = login_and_download(USERNAME, PASSWORD) 
    # upload to s3 or download locally
    save_emails(inbox, args.dir)


# handle args
def parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", help="input directory for email download")
    parser.add_argument("--upload", help="upload emails to AWS S3", action="store_true")
    args = parser.parse_args()
    # just making sure either --dir or --upload was used correctly
    if args.upload:
        return args
    elif args.dir:
        if os.path.isdir(args.dir):
            return args
        else:
            print("No such directory. Exiting...")
            quit()
    else:
        print("You must either enter --dir or --upload arguments")
        quit()


# connect to inbox and download messages
def login_and_download(username, password):
    auth = (username, password)
    inbox = Inbox(auth, getNow=False)
    # define 24-hour filter
    one_day_back = (datetime.utcnow() - timedelta(hours=24)).strftime("%Y-%m-%dT%H:%M:%SZ")
    inbox.setFilter("DateTimeReceived ge " + one_day_back)
    inbox.setOrderBy("DateTimeReceived desc")
    # download emails (using 1000 because defaults to only 10)
    try:
        result = inbox.getMessages(1000)
        # if nothing within last 24 hours, get recent 10 emails
        if len(inbox.messages) is 0:
            inbox.setFilter("")
            result = inbox.getMessages(1000)
    except requests.ConnectionError:
        print("Connection error during download. Exiting...")
        quit()
    # handle errors connecting/download
    if result:
        return inbox
    else:
        print("Error getting emails: " + inbox.errors)
        quit()

# iterate through emails and store them appropriately
def save_emails(inbox, path):
    index = 0
    for email in inbox.messages:
        # create filename for email
        filename = prepare_filename(index, email)
        index += 1
        # either upload to s3 or store locally
        if args.upload:
            upload(email, filename)
        else:
            save_locally(email, filename, path)
    print("Done. Stored " + str(index) + " emails.")


# create filename for email - (filename = [index]_[DateTimeReceived])
def prepare_filename(index, email):
    received = email.json["DateTimeReceived"].translate({ord(k):None for k in u"-:"})
    filename = str(index) + "_" + received
    return filename


# store email in specified directory
def save_locally(email, filename, path):
    with open(path + "/" + filename + ".txt", "w") as outfile:
         json.dump(email.json, outfile)


# upload email to s3
def upload(email, filename):
    filename += ".txt"
    client = boto3.client("s3")
    try:
        response = client.put_object(Bucket=BUCKET_NAME,
                                     Body=json.dumps(email.json),
                                     Key=filename)
    except client.exceptions.NoSuchBucket:
        print("Bucket doesn't exist. Exiting...")
        quit()
    except client.exceptions.EndpointConnectionError:
        print("Connection error during upload. Exiting...")
        quit()


if __name__ == "__main__":
    args = parse_arguments()
    main(args)
