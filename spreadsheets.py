import csv
from pickle import TRUE
import sqlite3
import re
import win32com.client as win32

# Connect to DB, in same folder as program
con = sqlite3.connect('allpos.db')

# Cursor for DB
cur = con.cursor()

# Get list of Vendors
result = cur.execute("SELECT DISTINCT(Vendor) FROM allpos").fetchall()
# Changes List of Tuples into Normal List
ven = list(sum(result,()))

# Count of Vendors 
ven_count = len(ven)
print(ven_count)


# GET POs for each vendor individually
for i in range(ven_count):
  first = ven[i]
  pos = cur.execute("SELECT Vendor, \"po number\", \"po type\", \"ship date\", \"newness flag\" FROM allpos WHERE Vendor = ?", [first]).fetchall()
   

#strips out non-alphanumeric characters for filename
  strip = re.sub(r'[^a-zA-Z0-9_ ]', '', ven[i])

  # Opens a File, writes header and all Vendor Rows to the file
  # Directory and file name is currently hardcoded and would need to be changed for other users / files
  # NOTE: Month names need to be manually updated - No longer matters
  open("C:\\Users\\asiegel\\Downloads\\Checkin Program\\csv\\%s Outstanding PO-Check-In.csv" % strip, "w").close()
  with open("C:\\Users\\asiegel\\Downloads\\Checkin Program\\csv\\%s Outstanding PO-Check-In.csv" % strip, "w+", newline='') as file:
       fieldnames = ['PO', 'PO Type', 'Ship Date', 'Newness flag', 'Vendor Confirmed Ship Date', 'Comments']
       writer = csv.DictWriter(file,fieldnames=fieldnames)
       writer.writeheader()

       # Write each PO
       for line in pos:
           print(line)
           row = {'PO':line[1], 'PO Type':line[2], 'Ship Date':line[3], 'Newness flag':line[4]}
           writer.writerow(row)


# TODO Sending Emails
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'dkuinlan@dwr.com' #TODO Create Email List
#mail.CC
mail.Subject = 'TEST CHECK-IN for %s' % ven[0]
mail.Body = 'This is a test email with %s\'s File. You will see an attachment too.' % ven[0]

# To attach a file to the email (optional):
attachment  = "C:\\Users\\asiegel\\Downloads\\Checkin Program\\csv\\%s Outstanding PO-Check-In.csv" % ven[0]
mail.Attachments.Add(attachment)

#mail.Display(True)
#mail.Send()