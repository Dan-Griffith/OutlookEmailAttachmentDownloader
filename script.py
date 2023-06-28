import os
import win32com.client

def download_email_attachments():
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        # Access default Inbox
        folder= inbox.Folders['FOLDER_NAME'] #Must be created as a subfolder in your inbox

        # Iterate over each email
        messages = folder.Items
        for message in messages:
            # If email has attachments
            if message.Attachments.Count > 0:
                print(f"Downloading attachments from email subject: {message.Subject}")

                # Iterate over each attachment
                for attachment in message.Attachments:
                    attachment_path = os.path.join("PATH_TO_DOWNLOAD", attachment.FileName)
                    # Save the attachment
                    attachment.SaveAsFile(attachment_path)
                    print(f"Attachment saved: {attachment_path}")

        print("Attachment download completed.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    download_email_attachments()
