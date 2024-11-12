from flask import Flask, request, render_template, redirect
import io
import pandas as pd     
import webbrowser as web 
from urllib.parse import quote 
import keyboard as k    
import time             

app = Flask(__name__)

# Function to send WhatsApp messages
def send_whatsapp(excel_file, message_template):
    # Read the Excel file into a pandas DataFrame from the in-memory file object
    df = pd.read_excel(excel_file, dtype={"Contact": str})
    names = df['Name'].values
    contacts = df['Contact'].values

    web.open("https://web.whatsapp.com")
    time.sleep(20)  # Adjust this to allow time for QR code scanning if not already logged in

    # Loop over each contact and send a message in the same tab
    for name, contact in zip(names, contacts):
        try:
            # Ensure that the name is correctly replaced in the message template
            personalized_message = message_template.format(name=name)
        except KeyError as e:
            print(f"Error in template formatting for {name}: {e}")
            continue  # Skip this contact if the template has an issue

        # Construct the WhatsApp Web URL for the given phone number and message
        whatsapp_url = f"https://web.whatsapp.com/send?phone={contact}&text={quote(personalized_message)}"
        
        # Navigate to the new URL by typing it into the address bar in the same tab
        k.press_and_release('ctrl+l')  # Focus the browser address bar
        time.sleep(1)
        k.write(whatsapp_url)
        k.press_and_release('enter')
        
        # Wait for the chat to load
        time.sleep(15)  # Adjust if loading time varies
        
        # Automatically press Enter to send the message
        k.press_and_release('enter')
        time.sleep(2)  # Short delay to ensure message is sent
        
        print(f"Message sent to {name}")

    print("All messages sent successfully!")

# Route for the index page to upload files
@app.route('/')
def index():
    return render_template('index.html')

# Route for file upload and message sending
@app.route('/send', methods=['POST'])
def send():
    # Check if the Excel file is part of the request
    if 'excel_file' not in request.files:
        return redirect(request.url)
    
    excel_file = request.files['excel_file']
    message_template = request.form['message_template']  # Get the message template from the form input

    if excel_file and message_template:
        # Process the files directly without saving them
        send_whatsapp(excel_file, message_template)

        return "All messages sent successfully!"

    return "Failed to upload files or missing message template."

if __name__ == '__main__':
    app.run(debug=True)
