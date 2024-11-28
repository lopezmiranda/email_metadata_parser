import os
import csv
import email
from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup
from datetime import datetime
import shutil

def move_file(src, dest):
    # Ensure the source file exists
    if not os.path.isfile(src):
        return f"Error: The source file '{src}' does not exist."
    
    # Ensure the destination directory exists
    if not os.path.isdir(dest):
        return f"Error: The destination directory '{dest}' does not exist."
    
    # Move the file
    try:
        shutil.move(src, dest)
        return f"File successfully moved from '{src}' to '{dest}'."
    except Exception as e:
        return f"Error: {e}"

# Generate current date and time
def get_current_datetime_string():
    # Get the current date and time
    now = datetime.now()
    
    # Format the string as 'YYYYMMDDHH' (Year, Month, Day, Hour)
    return now.strftime("%Y%m%d%H")

# Function to extract links from email body text (ignore mailto: links)
def extract_links_from_body(body):
    links = []
    soup = BeautifulSoup(body, "html.parser")
    # Find all links in the body
    for anchor in soup.find_all('a', href=True):
        link = anchor['href']
        # Exclude mailto links
        if not link.startswith("mailto:"):
            links.append(link)
    return links

# Function to determine if the link is from a known File Storage service
def is_filestorage_link(link):
    known_filestorage_services = ["drive.google.com", "box.com", "dropbox.com"]
    return any(service in link for service in known_filestorage_services)

# Parse .eml file
def parse_eml(file_path):
    # Read .eml file
    with open(file_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)

    # Extract sender, recipients, and date
    message_id = msg.get('Message-ID')
    sender = msg.get('From')
    subject = msg.get('Subject')
    to = msg.get('To')
    cc = msg.get('Cc')
    date = msg.get('Date')

    # EXTRACT ATTACHMENTS
    attachments = []

    # Helper function to extract attachments recursively from multipart parts
    def extract_attachments(part):
        # If the part has an attachment disposition, save the filename
        content_disposition = part.get("Content-Disposition", "")
        if content_disposition and "attachment" in content_disposition:
            filename = part.get_filename()
            if filename and filename != "smime.p7s":
                attachments.append(filename)

        # If the part is multipart, recurse through its subparts
        if part.is_multipart():
            for subpart in part.iter_parts():
                extract_attachments(subpart)

    # Start extraction from the root email message
    extract_attachments(msg)

    # Extract email body (text or HTML)
    body = ""
    for part in msg.iter_parts():
        if part.get_content_type() == 'text/plain':
            body += part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')
        elif part.get_content_type() == 'text/html':
            body += part.get_payload(decode=True).decode(part.get_content_charset(), errors='ignore')

    # Extract links from email body
    links = extract_links_from_body(body)
    file_links = []
    for link in links:
        if (is_filestorage_link(link)):
            file_links.append(link)

    return {
        "message_id": message_id,
        "sender": sender,
        "to": to,
        "cc": cc,
        "date": date,
        "subject": subject,
        "attachment_count": len(attachments),
        "attachment_names": ', '.join(attachments),
        "link_count": len(file_links),
        "links": ', '.join(file_links)
    }


# Write to CSV file
def write_to_csv(data, csv_path):
    # Check if the CSV file exists
    file_exists = os.path.exists(csv_path)

    # Open the file in append mode
    with open(csv_path, mode='a', newline='', encoding='utf-8') as file:
        fieldnames = ["message_id", "sender", "to", "cc", "date", "subject", "attachment_count", "attachment_names", "link_count", "links"]
        writer = csv.DictWriter(file, fieldnames=fieldnames)

        # If the file doesn't exist, write the header row
        if not file_exists:
            writer.writeheader()

        # Write the data row
        writer.writerow(data)

def print_email_content(email_object):
    # Helper function to return '-' if a value is empty or falsy
    def get_field_value(field):
        return field if field else '-'
    
    # Multiline string with placeholders for values from email_object
    message = f"""
    Estimada(o), el presente mensaje es para constatar la recepción de los datos que comparte con la GCRMN-ETPN. Desglosamos los detalles del envío:

    Asunto: {get_field_value(email_object.get("subject", ""))}
    Emisor(a): {get_field_value(email_object.get("sender", ""))}
    Receptores: {get_field_value(email_object.get("to", ""))}, {get_field_value(email_object.get("cc", ""))}
    Enlaces encontrados en el correo: {get_field_value(email_object.get("links", ""))}
    Archivos adjuntos: {get_field_value(email_object.get("attachment_names", ""))}
    Bases de datos en el contenido:
    """
    
    # Print the formatted message
    print(message)

# Main function
def process_eml_files(input_dir, output_dir, csv_path):
    for filename in os.listdir(input_dir):
        if filename.endswith(".eml"):
            file_path = os.path.join(input_dir, filename)
            try:
                email_data = parse_eml(file_path)
                write_to_csv(email_data, csv_path)
                move_file(file_path, output_dir)
                print(f"Processed {file_path}")
                print_email_content(email_data)
            except Exception as e:
                print(f"Error processing {file_path}: {e}")

# Specify the input directory containing .eml files and the CSV file path
input_directory = "./not_processed"  # Path to directory with unprocessed .eml files
output_directory = "./processed" # Path to store processed .eml files
csv_file_path = f"./data_metadata/email_metadata_procesados.csv"  # Path to CSV output

# Process the .eml files and write metadata to CSV
process_eml_files(input_directory, output_directory, csv_file_path)

