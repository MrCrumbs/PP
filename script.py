from O365 import Inbox
from datetime import datetime, timedelta
import argparse, json, re, requests, os, boto3, botocore, csv, pandas
USERNAME = "qa.ex@office365.ecknhhk.xyz"
PASSWORD = "ew68I7W52p*W"
BUCKET_NAME = "interview-exercises"

# main flow
def main(args):
    # connect to inbox and download messages
    inbox = login_and_download(USERNAME, PASSWORD, args.report)
    # upload to s3 or download locally
    save_emails(inbox, args.dir, args.report)


# handle args
def parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", help="input directory for email download")
    parser.add_argument("--upload", help="upload emails to AWS S3", action="store_true")
    parser.add_argument("--report", help="generate email report", action="store_true")
    args = parser.parse_args()
    # just making sure either --dir or --upload was used correctly
    if args.upload:
        return args
    elif args.dir:
        if os.path.isdir(args.dir):
            return args
        else:
            print("No such directory. Exiting...")
    else:
        print("You must either enter --dir or --upload arguments")
    quit()


# connect to inbox and download messages
def login_and_download(username, password, report):
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
    # generate report of downloaded emails, for now without s3 link
    if report:
        generate_report(inbox.messages)
    # handle errors connecting/download
    if result:
        return inbox
    else:
        print("Error getting emails: " + inbox.errors)
        quit()

# iterate through emails and store them appropriately
def save_emails(inbox, path, report):
    index = 0
    emails_stored = 0
    for email in inbox.messages:
        # create filename for email
        filename = prepare_filename(index, email)
        index += 1
        # either upload to s3 or store locally
        try:
            if args.upload:
                upload(email, filename, report, index)
            else:
                save_locally(email, filename, path)
            emails_stored += 1
        except:
            pass
    print("Done. Stored " + str(emails_stored) + " emails.")


# create filename for email - (filename = [index]_[DateTimeReceived])
def prepare_filename(index, email):
    received = email.json["DateTimeReceived"].translate({ord(k):None for k in u"-:"})
    filename = str(index) + "_" + received
    return filename


# store email in specified directory
def save_locally(email, filename, path):
    try:
        with open(path + "/" + filename + " .txt", "w") as outfile:
            json.dump(email.json, outfile)
    except Exception as e:
        print("Error writing " + filename + " to file.\n" + 
              "Error message: " + str(e) + "\nMoving on...")
        raise


# upload email to s3
def upload(email, filename, report, index):
    filename += ".txt"
    client = boto3.client("s3")
    # write email to s3, and update report with url
    try:
        response = client.put_object(Bucket=BUCKET_NAME,
                                     Body=json.dumps(email.json),
                                     Key=filename)
        if report:
            url = client.generate_presigned_url("get_object", 
                                                Params = {"Bucket": BUCKET_NAME, 
                                                          "Key": filename})
            update_url_in_report(url, index)
    except client.exceptions.NoSuchBucket:
        print("Bucket doesn't exist. Exiting...")
        quit()
    except (botocore.exceptions.EndpointConnectionError,
            botocore.vendored.requests.exceptions.ConnectionError):
        print("Connection error during upload of " + filename + ". Moving on...")
        raise

# generate email report, for now without s3 link
def generate_report(emails):
    writer = csv.writer(open("report.csv", "w"))
    writer.writerow(["Time", "From", "To", "Subject", "S3 Link"])
    for email in emails:
        writer.writerow([email.json["DateTimeReceived"].encode('utf-8'), 
                        email.json["Sender"]["EmailAddress"]["Address"].encode('utf-8'), 
                        email.json["ToRecipients"][0]["EmailAddress"]["Address"].encode('utf-8'), 
                        email.json["Subject"].encode('utf-8'),
                        "Not uploaded".encode('utf-8')])


# when email is uploaded to s3 - update report with link
def update_url_in_report(url, index):
    df = pandas.read_csv("report.csv")
    df.iat[index-1, 4] = url
    df.to_csv("report.csv", index=False)
    
if __name__ == "__main__":
    args = parse_arguments()
    main(args)
