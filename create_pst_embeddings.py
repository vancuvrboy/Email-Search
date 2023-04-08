# This script creates embeddings from an Outlook PST file.
# It uses the python-pst library to read the PST file.
# The command to install the python-pst library is "pip install Aspose.Email-for-Python-via-NET".

import os
import sys
import json
import openpyxl
import csv
import re
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import joblib

#from aspose.email.storage.pst import PersonalStorage

from libratom.lib.pff import PffArchive
from email import generator
from pathlib import Path
from base64 import urlsafe_b64decode, urlsafe_b64encode
from sklearn.tree import DecisionTreeClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score,classification_report,confusion_matrix
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.naive_bayes import MultinomialNB, GaussianNB
from sklearn import svm
from sklearn.svm import SVC
from sklearn.model_selection import GridSearchCV
from sklearn.feature_extraction.text import TfidfVectorizer


def get_training_emails_from_PST():
    #read emails from PST file and save in an excel file
    archive = PffArchive("test.pst")
    #eml_out = Path(Path.cwd() / "emls")

    #if not eml_out.exists():
        #eml_out.mkdir()

    for f in archive.folders():
        if f.name == "Inbox":
            folder = f
            break

    # define tests to categorize messages
    subject_calendar = ["Accepted:",
                                "Declined:",
                                "Canceled:",
                                "Updated:",
                                "Tentative:",
    ]
    body_calendar_rejects=["When:",
    ]                   
    subject_rejects = ["On Demand Webinar Request",
                        "Contributor Add-In Auto Trial Activation",
                        "Colligo Contributor Trial for",
                        "Out of Office AutoReply",
    ]
    from_rejects=["<contributor@colligo.com>",
                "element 5",
                "<sales@colligo.com>",
                "Plaxo",
                "Vancouver Enterprise Forum Group Members",
                "SharePoint Magazine User Group Group Members",
                "Colligo for SharePoint User Group",
                "WTIA Events",
                "Central Desktop Digest",
                "SpyFu News",
                "Colligo Activation",
                "IT Security Bulletin",
                "Orbitz",
                "Tripit",
                "irewards",
                "Dell | Small Business",
                "Mileage Plus Partners",
                "<partners@colligo.com>",
                "Lenovo",
                "SearchWinIT.com",
                "WordPress",
                "Facebook",
                "Twitter",
                "LinkedIn",
                "Google+",
                "YouTube",
                "Microsoft",
                "LeadForce1 Team",
                "echannelline@echannelline.com",
                "Sean Elbe",
                "The 451 Group",
                "Zona Ixyhjl",
                "TigerDirect.com",
                "Colligo Networks",
                "DangleTech",
                "Tim Bramwell",
                "SearchNetworking.com",
                "Gartner Events",
                "\"Linux & Open Source eSeminar",
                "Ferris News Service",
                "Virtualization Bulletin",
                "Software Developer Bulletin",
                "Sophie Heywood",
                "\"Heywood, Sophie",
                "Alon Newton",
                "Martin Muldoon",
                "IDG Connect",
                "Frost & Sullivan",
                "InteleTravel.com",
                "Ferris News Service",
                "\"onDemand eSeminars",
                "WINBC",
                "British Columbia Innovation Council",
                "<gotowebinar@citrixonline.com>",
                "Costco News",
                "Smile City Square",
                "Office and SharePoint Pro UPDATE",
                "Vancouver MLS",
                "Arden Pennell",
                "Air Canada webSaver",
                "<gotowebinar@citrixonline.com>",
                "\"Wednesday's Linux & Open Source eSeminar",
                "Aditya Sanyal",
                "The Stevie Awards",
                "Ed Chyzowski",
    ]
    from_marketing = ["Melanie Aizer", 
                        "Randy Halischuk",
                        "Genese Castonguay",
                        "Asa Zanatta",
                        "\"Asa Zanatta",
                        "Sarah Johnson",
                        "<azanatta@spiritcommunications.ca>",
                        "Michael Sampson",
    ]
    from_sales = ["Ed Kaczor",
                    "Braeden Calyniuk",
    ]
    from_finance = ["Siena Sayasaya",
                    "Simon Kok",
    ]
    from_networking = []
    from_HR = ["David via Central Desktop",
                "Don Safnuk",
                "Mike Chew",
    ]
    from_personal = ["Kevin Leong",
                        "\"Bob @ WPX",
                        "Judi Richardson",
                        "David MacLaren",
                        "Suzie Dean",
                        "Carrothers",
                        "Bruce Dean",
                        "<suzie@northamgroup.com>",
                        "Paul Carrothers",
                        "Colleen Vickruck",
                        "Paul Manning",
                        "<randy.garg@ca.pwc.com>",
                        "Saxon Shuttleworth",
    ]
    from_vc = ["Thomas Bailey",
                "\"Robert D. Valdez",
                "Sarah Tavel",
    ]
    from_eng = ["Dave Foster",
                "Andrew Block",
    ]

    n_messages = 0
    rejected_messages = 0
    for message in folder.sub_messages:
        # subject is type str
        # print(message.subject)   
        # sender_name is type str 
        # print(message.sender_name)
        # transport_headers is type str
        from_name = None
        to = None
        cc = None
        bcc = None
        subject = None
        date = None
        message_fields = {}

        # initially categorize the email as undefined
        message_fields['CLASS NAME'] = "U"
        # extract message properties from transport_headers
        for header in message.transport_headers.splitlines():
            if header.lstrip().startswith("From:"):
                from_name = header.lstrip().split("From:")[1].strip()
                message_fields['from_name'] = from_name
                # categorize the email by who it is from
                if [True for reject in from_rejects if from_name.startswith(reject)]: 
                    message_fields['CLASS NAME'] = "R"
                elif [True for m in from_marketing if from_name.startswith(m)]:
                    message_fields['CLASS NAME'] = "M"
                elif [True for s in from_sales if from_name.startswith(s)]:
                    message_fields['CLASS NAME'] = "S"
                elif [True for f in from_finance if from_name.startswith(f)]:
                    message_fields['CLASS NAME'] = "F"
                elif [True for n in from_networking if from_name.startswith(n)]:
                    message_fields['CLASS NAME'] = "N"
                elif [True for h in from_HR if from_name.startswith(h)]:
                    message_fields['CLASS NAME'] = "H"
                elif [True for p in from_personal if from_name.startswith(p)]: 
                    message_fields['CLASS NAME'] = "P"
                elif [True for v in from_vc if from_name.startswith(v)]:
                    message_fields['CLASS NAME'] = "V"
                elif [True for e in from_eng if from_name.startswith(e)]:
                    message_fields['CLASS NAME'] = "E"
            if header.lstrip().startswith("To:"):
                to = header.lstrip().split("To:")[1].strip()
                message_fields['to'] = to
            if header.lstrip().startswith("CC:"):
                cc = header.lstrip().split("CC:")[1].strip()
                message_fields['cc'] = cc
            if header.lstrip().startswith("BCC:"):
                bcc = header.lstrip().split("BCC:")[1].strip()
                message_fields['bcc'] = bcc
            if header.lstrip().startswith("Subject:"):
                subject = header.lstrip().split("Subject:")[1].strip()
                # remove pipe characters from subject
                message_fields['subject'] = subject.replace('|','')
                # categorize the email by the subject
                if [True for reject in subject_rejects if subject.startswith(reject)]:
                    message_fields['CLASS NAME'] = "R"
                elif [True for c in subject_calendar if subject.startswith(c)]:
                    message_fields['CLASS NAME'] = "C"
            if header.lstrip().startswith("Date:"):
                date = header.lstrip().split("Date:")[1].strip()
                message_fields['date'] = date
        # extract message body
        # plain_text_body is type bytes
        if message.plain_text_body != None:
            sbod = message.plain_text_body.decode('utf-8').lstrip()
            # remove pipe characters, commas and html tags from body
            # limit to 30,000 characters (approximately 8000 tokens)
            message_fields['body'] = re.sub('<[^<]+?>', '',sbod.replace('|',''))[0:30000]
            # categorize the email by the body
            if [True for c in body_calendar_rejects if sbod.startswith(c)]:
                message_fields['CLASS NAME'] = "C"
        # html_body is type bytes (presumed)
        #if message.html_body != None:
            #hbod = message.html_body.decode('utf-8').lstrip()
            #message_fields['html_body'] = hbod
        # delivery_time is type datetime    
        #if message.delivery_time != None:
            #message_fields['delivery_time'] = message.delivery_time 

        if message.number_of_attachments > 0:
            # read from attachment 1
            attach_size = message.get_attachment(0).get_size()
            message_fields['attach_size'] = attach_size
            #attachment_content = (message.get_attachment(0).read_buffer(attach_size)).decode('ascii', errors='ignore')
            #message_fields['attachment_content'] = attachment_content

        # if the message is to be rejected, don't write it to the csv file
        if message_fields['CLASS NAME'] == "R" or message_fields['CLASS NAME'] == "C":
            rejected_messages += 1
            continue

        # write the message to the csv file
        try:
            with open('training_set.csv', 'a', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['CLASS NAME', 'from_name', 'to', 'cc', 'date', 'subject', 'attach_size', 'body', 'attachment_content']
                # docs on csv https://docs.python.org/3/library/csv.html#dialects-and-formatting-parameters
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writerow(message_fields)
        except Exception as e:
            print(f"Error writing to csv file: {e}")
            print (f"message_fields: {message_fields}")
            print(f"processed {n_messages} messages when error happened")
            continue

        # debug emails that don't write properly to the csv file
        #if n_messages >= 660 and n_messages < 675:
            #print(f"message number: {n_messages}")
            #print(f"message_fields: {message_fields}")
            #print(50*"=")
        

        n_messages += 1
        #print(f"Processed {n_messages} messages")
        if n_messages > 1000:
            break

    print(f"Done! {n_messages-1} messages processed and {rejected_messages} rejected")
    return

def load_data():
    # Load data from a CSV file and return X and y using pandas data frame
    fieldnames = ['CLASS NAME', 'from_name', 'to', 'cc', 'date', 'subject', 'attach_size', 'body', 'attachment_content']
    #fieldnames = ['CLASS NAME', 'from_name', 'to', 'cc', 'date', 'subject']
    #df = pd.read_csv('classified_training_set.csv', names = fieldnames, sep=',', encoding='utf-8')
    df = pd.read_csv('classified_training_set.csv', names = fieldnames, sep=',')
    X = "From: "+df['from_name']+"\n"+"To: "+df['to']+"\n"+"Subject: "+df['subject']+"\n"+"Body: "+df['body']

    #X = df.drop('CLASS NAME', axis=1)
    #X = df['body']
    y = df['CLASS NAME']
    return X.values.astype('U'), y.values.astype('U')


def train_classifier_grid_search():
    X, y = load_data()
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

    # Train the classification model

    #clf = DecisionTreeClassifier(criterion="gini", random_state=42, max_depth=3, min_samples_leaf=5) 
    #clf.fit(X_train, y_train)

    count_vector = CountVectorizer(stop_words='english', max_features=500)
    extracted_features = count_vector.fit_transform(X_train)

    #print(f"Extracted features:\n{extracted_features}")

    print("Training the model...")

    tuned_parameters = {'kernel': ['rbf','linear'], 'gamma': [1e-3, 1e-4],'C': [1, 10, 100, 1000]}
    model = GridSearchCV(svm.SVC(), tuned_parameters)
    model.fit(extracted_features,y_train)

    # save the model to disk
    filename = 'finalized_model.sav'
    joblib.dump(model, filename)


    print("Model Trained Successfully!")

    # Test the accuracy of the model
    print("Testing the model...")
    #y_predict = clf.predict(X_test)
    score = model.score(count_vector.transform(X_test),y_test)*100
    print(f"Accuracy: {score}")

    # Save the model to a file
    # joblib.dump(clf, 'email_classifier.pkl')



    # Use the model to classify new emails 
    #emails = load_new_emails()
    #predictions = clf.predict(emails)

def train_classifier_SVM():
    X, y = load_data()
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

    # Train the classification model

    #clf = DecisionTreeClassifier(criterion="gini", random_state=42, max_depth=3, min_samples_leaf=5) 
    #clf.fit(X_train, y_train)

    vectorizer = TfidfVectorizer(stop_words='english')

    vec_file = 'vectorizer.pkl'
    joblib.dump(vectorizer, vec_file)

    X_train = vectorizer.fit_transform(X_train)




    X_test = vectorizer.transform(X_test)

    #print(f"Extracted features:\n{extracted_features}")

    print("Training the model...")

    model = SVC(kernel='linear', C=1.0, random_state=42)
    model.fit(X_train,y_train)

    # save the model to disk
    filename = 'svm_model.pkl'
    joblib.dump(model, filename)


    print("Model Trained Successfully!")

    # Test the accuracy of the model
    print("Testing the model...")
    #y_predict = clf.predict(X_test)
    score = model.score(X_test,y_test)*100
    print(f"Accuracy: {score}")

    # Save the model to a file
    # joblib.dump(clf, 'email_classifier.pkl')



    # Use the model to classify new emails 
    #emails = load_new_emails()
    #predictions = clf.predict(emails)

def classify_emails():
    # Use the model to classify new emails 
    emails, email_data_frame = load_new_emails()
    #predictions = clf.predict(emails)
    filename = 'finalized_model.sav'
    model = joblib.load(filename)
    count_vector = CountVectorizer(stop_words='english', max_features=500)
    extracted_features = count_vector.fit_transform(emails)
    predictions = model.predict(extracted_features)
    print(f"predictions: {predictions}")
    write_predictions_to_csv(predictions, email_data_frame)
    return predictions

def classify_emails_SVM():
    # Use the model to classify new emails 
    emails, email_data_frame = load_new_emails()
    #predictions = clf.predict(emails)
    filename = 'svm_model.pkl'
    model = joblib.load(filename)

    # load vectorizer from file
    vec_file = 'vectorizer.pkl'
    vectorizer = joblib.load(vec_file)

    #vectorizer = TfidfVectorizer(stop_words='english')
    #text_vector = vectorizer.fit_transform(emails)
    text_vector = vectorizer.fit_transform(emails)

    #use SVM model to predict the class

    predictions = model.predict(text_vector)
    print(f"predictions: {predictions}")
    write_predictions_to_csv(predictions, email_data_frame)
    return predictions

def write_predictions_to_csv(predictions, edf):
    # Write the predictions to a CSV file
    fieldnames = ['CLASS NAME', 'from_name', 'to', 'cc', 'date', 'subject', 'body']

    edf = edf.drop(['attach_size', 'attachment_content'], axis=1)
    print(edf['from_name'])

    try:
        with open('predictions.csv', 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            i = 0
            for prediction in predictions:
                writer.writerow({'CLASS NAME': prediction,'from_name' : edf['from_name'][i],'to' : edf['to'][i],'cc' : edf['cc'][i],'date' : edf['date'][i],'subject' : edf['subject'][i],'body' : edf['body'][i]})
                i += 1
    except Exception as e:
        print(f"Error writing to csv file: {e}")
    return

def load_new_emails():
    # Load new emails, e.g. from a CSV file
    # and return the data as X
    fieldnames = ['from_name', 'to', 'cc', 'date', 'subject', 'attach_size', 'body', 'attachment_content']
    df = pd.read_csv('test_data.csv', names = fieldnames, sep=',')
    X = "From: "+df['from_name']+"\n"+"To: "+df['to']+"\n"+"Subject: "+df['subject']+"\n"+"Body: "+df['body']
    return X.values.astype('U'), df

if __name__ == "__main__":
    #get_training_emails_from_PST()
    #train_classifier_grid_search()
    #train_classifier_SVM()
    #load_data()
    classify_emails()
    #classify_emails_SVM()
