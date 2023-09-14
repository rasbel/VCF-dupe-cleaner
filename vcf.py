import vobject
import os
import re
import win32com.client
#import independentsoft.msg
from independentsoft.msg import Message
#from msgpy.reader import MessageReader

# Replace 'your_file.vcf' with the path to your VCF file
vcf_file = 'vcf_file.vcf'

# Read and parse the VCF file
with open(vcf_file, 'r', encoding='utf-8') as file:
    vcf_data = file.read()

# Parse the VCF data into a list of vCards
vcards = list(vobject.readComponents(vcf_data))

# Create a set to track unique vCard representations
unique_vcards = set()

# Create a list to store unique vCards
unique_vcard_list = []

# Iterate through the vCards and remove duplicates
for vcard in vcards:
    # Convert each vCard to a string to check for uniqueness
    vcard_str = vcard.serialize()
    
    # Check if this vCard representation is unique
    if vcard_str not in unique_vcards:
        unique_vcards.add(vcard_str)
        unique_vcard_list.append(vcard)

# Print unique contact information
for vcard in unique_vcard_list:
    print("Name:", vcard.fn.value)
    #print("Email:", vcard.email.value)
    #print("Phone:", vcard.tel.value)
    print()
# Create a new VCF file and write the unique vCards to it
new_vcf_file = 'new_unique_contacts.vcf'

# Open the new VCF file for writing
with open(new_vcf_file, 'w', encoding='utf-8') as file:
    # Iterate through the unique vCards and serialize them back to VCF format
    for vcard in unique_vcard_list:
        vcard_str = vcard.serialize()
        file.write(vcard_str)
        file.write('\n')  # Add a newline separator between vCards

print(f"Unique vCards saved to '{new_vcf_file}'.")

# Create an Outlook application object
outlook = win32com.client.Dispatch("Outlook.Application")

# Access the default Contacts folder
namespace = outlook.GetNamespace("MAPI")
contacts_folder = namespace.GetDefaultFolder(10)  # 10 represents the Contacts folder

# Iterate through the contacts and extract information
for contact_item in contacts_folder.Items:
    print("Name:", contact_item.FullName)
    print("Email:", contact_item.Email1Address)
    print("Phone:", contact_item.BusinessTelephoneNumber)
    print()

# Release Outlook resources
outlook.Quit()
Directory containing .msg files
msg_directory = r"C:\Users\aassefa\OneDrive - cabreracapital.com\Desktop\vcfcards\Robert's Cielo Contacts"

# Dictionary to store unique contact information (email addresses)
unique_contacts = {}

# Iterate through .msg files in the directory
for filename in os.listdir(msg_directory):
    if filename.endswith('.msg'):
        msg_filepath = os.path.join(msg_directory, filename)
        
        # Read the .msg file and extract contact information (email addresses)
        with open(msg_filepath, 'r', encoding='utf-8') as msg_file:
            try:
                msg_content = msg_file.read()
            except:
                print("Message couldn't be read")
                continue
            # Use regular expression to find email addresses (adjust the pattern as needed)
            email_addresses = re.findall(r'\S+@\S+', msg_content)

            # Add email addresses to the dictionary to identify duplicates
            for email in email_addresses:
                unique_contacts.setdefault(email, []).append(msg_filepath)

# Iterate through unique contacts and keep only one .msg file per contact
for email, msg_files in unique_contacts.items():
    if len(msg_files) > 1:
        # Decide which .msg file to keep (e.g., based on file creation date or other criteria)
        # Replace the logic here to select the file to keep as needed
        file_to_keep = msg_files[0]

        # Delete duplicate .msg files
        for duplicate_file in msg_files[1:]:
            os.remove(duplicate_file)

print("Duplicate .msg files removed.")
