import json
import html
import re
from docx import Document
from datetime import datetime

def load_conversation(json_file):
    """Load the Skype conversation JSON file."""
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

def clean_content(text):
    """
    Clean the message content by:
    1. Unescaping HTML entities.
    2. Removing XML/HTML tags like <e_m ...></e_m>.
    """
    # Unescape HTML entities (e.g., &#x27; becomes ', &amp; becomes &)
    text = html.unescape(text)
    # Remove tags of the form <e_m ...></e_m> or self-closing <e_m .../>
    text = re.sub(r'<e_m\b[^>]*>(.*?)</e_m>', '', text)
    text = re.sub(r'<e_m\b[^>]*/>', '', text)
    return text

def format_date(iso_date):
    """
    Convert an ISO formatted date string into a nicer format.
    For example: '2025-03-19T21:15:44.097Z' -> 'March 19, 2025 09:15:44 PM'
    """
    try:
        # Remove the trailing 'Z' if present.
        if iso_date.endswith("Z"):
            iso_date = iso_date[:-1]
        dt = datetime.fromisoformat(iso_date)
        # Format in 12-hour time with AM/PM.
        return dt.strftime("%B %d, %Y %I:%M:%S %p")
    except Exception:
        return iso_date  # Return original if parsing fails.

def filter_messages(data, username, first_word):
    """
    Filter messages from all conversations in the data.
    
    A message will be selected if the sender (ignoring any prefix like '8:')
    matches the provided username and the message's first word exactly matches.
    """
    filtered = []
    conversations = data.get("conversations", [])
    for convo in conversations:
        messages = convo.get("MessageList", [])
        for message in messages:
            sender = message.get("from", "")
            # Remove prefix (like "8:") if present.
            if ":" in sender:
                sender = sender.split(":", 1)[1]
            content = message.get("content", "")
            # Clean the message content.
            content = clean_content(content)
            words = content.strip().split()
            if not words:
                continue  # Skip empty messages.
            if sender == username and words[0] == first_word:
                # Format the date for display.
                iso_date = message.get("originalarrivaltime", "Unknown Date")
                formatted_date = format_date(iso_date)
                filtered.append({
                    "date": formatted_date,
                    "content": content
                })
    return filtered

def save_to_word(filtered_messages, output_file, username, first_word):
    """
    Save filtered messages to a Word document.
    
    The document includes a header with the username and search criteria, 
    and each message is separated by a heading that includes the formatted message date.
    """
    doc = Document()
    # Add a header with the username and search criteria.
    doc.add_heading(f"Messages by {username} ", level=1)
    
    for i, msg in enumerate(filtered_messages, start=1):
        # Display the formatted date of the current message in the heading.
        doc.add_heading(f"Message {i} - {msg['date']}", level=2)
        doc.add_paragraph("Message content:")
        doc.add_paragraph(msg["content"])
        doc.add_paragraph("-" * 40)  # Separator line.
    
    doc.save(output_file)
    print(f"Document saved as {output_file}")

def main():
    json_file = "skype_conversation.json"  # Update the path if necessary.
    try:
        data = load_conversation(json_file)
    except Exception as e:
        print(f"Error reading JSON file: {e}")
        return

    # Get user input.
    username = input("Enter the username to search for: ").strip()
    first_word = input("Enter the first word to match: ").strip()
    
    filtered_messages = filter_messages(data, username, first_word)
    if not filtered_messages:
        print("No messages matched the given criteria.")
        return

    output_file = (f"{username}_messages.docx")
    save_to_word(filtered_messages, output_file, username, first_word)

if __name__ == "__main__":
    main()
