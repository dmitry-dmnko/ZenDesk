import win32com.client


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
root_folder = outlook.Folders.Item(1)

indir_Mark = "C:\\Users\\DmitryDmytrenko\\Documents\\Technology\\DD\\Zendesk\\Py app\\Files From Brad Mark"
indir_Merch = "C:\\Users\\DmitryDmytrenko\\Documents\\Technology\\DD\\Zendesk\\Py app\\Files From Brad Merch"


def get_mark_mail():
    subfolder = root_folder.Folders['Inbox'].Folders['ZD Marketing']
    messages = subfolder.Items
    message = messages.GetFirst()
    subject = message.Subject
    for m in messages:
        try:
            attachments = message.Attachments
            attachment1 = attachments.Item(1)
            attachment1.SaveASFile(indir_Mark + '\\' + str(attachment1))
            attachment2 = attachments.Item(2)
            attachment2.SaveASFile(indir_Mark + '\\' + str(attachment2))
            attachment3 = attachments.Item(3)
            attachment3.SaveASFile(indir_Mark + '\\' + str(attachment3))
            attachment4 = attachments.Item(4)
            attachment4.SaveASFile(indir_Mark + '\\' + str(attachment4))
            message = messages.GetNext()

        except:
            message = messages.GetNext()


def get_merch_mail():
    subfolder = root_folder.Folders['Inbox'].Folders['ZD Merchandising']
    messages = subfolder.Items
    message = messages.GetFirst()
    subject = message.Subject
    for m in messages:
        try:
            attachments = message.Attachments
            attachment1 = attachments.Item(1)
            attachment1.SaveASFile(indir_Merch + '\\' + str(attachment1))
            attachment2 = attachments.Item(2)
            attachment2.SaveASFile(indir_Merch + '\\' + str(attachment2))
            attachment3 = attachments.Item(3)
            attachment3.SaveASFile(indir_Merch + '\\' + str(attachment3))
            attachment4 = attachments.Item(4)
            attachment4.SaveASFile(indir_Merch + '\\' + str(attachment4))
            message = messages.GetNext()

        except:
            message = messages.GetNext()


get_mark_mail()
get_merch_mail()
