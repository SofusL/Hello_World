import win32com.client as win32

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

mailItem = olApp.CreateItem(0)
mailItem.BodyFormat = 1
mailItem.Sensitivity  = 2

# ---------------------------------------------------------------------------------------------- #

# - DATA FORMAT (VAT NUMBER, CLIENT NAME, FOLDER NAME, EMAIL MAIN, EMAIL CC)
DATA = [
('1','Sofus','SofusFolder'),
('2','Oliver','OliverFolder'),
('3','Villads','VilladsFolder')
  ]

DATE = ("2023 01") # - ENTER THE CURRENT DATE (YYYY MM)

# ---------------------------------------------------------------------------------------------- #

def find_path(name, path):
  for root, dirs, files in os.walk(path):
    if name in dirs or name in files:
      return root
  return None

CLIENT = input("What is the VAT number of the client you want to send this to?")

found = False
for record in DATA:
  if CLIENT == record[0]:
    print("Sender e-mail til", record[1])
    found = True
    root_dir = "C:\\Users\\soffe\\Desktop\\" + record[2]
    target_name = DATE + ' Tolddeklarationsoversigt.txt'
    path = find_path(target_name, root_dir)
    if path:
      print(f"Found at: {path}")
      mailItem.Subject = "Tolddeklarationsoversigt " + DATE + " ("+record[1]+")"
      mail.Attachments.Add(path)
      mailItem.Display

    else:
      print(f"{target_name} not found")
      print(root_dir)
    break

if not found:
  print(CLIENT + " not found")

# ---------------------------------------------------------------------------------------------- #


