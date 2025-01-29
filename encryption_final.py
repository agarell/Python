from cryptography.fernet import Fernet
import base64
import os
import sys

# creates encryption key
def generate_key():
    return Fernet.generate_key()

# loads encryption key to be used
def load_key(key):
    return Fernet(key)

# function to encrypt files, writing to them in binary for correct encrpytion (wb)
def encrypt_file(file_path, key):
    fernet = load_key(key)
    with open(file_path, 'rb') as file:
        original = file.read()
    encrypted = fernet.encrypt(original)
    with open(file_path, 'wb') as encrypted_file:
        encrypted_file.write(encrypted)

# function to decrypt files
def try_decrypt_file(file_path, key):
    fernet = load_key(key)
    with open(file_path, 'rb') as file:
        encrypted = file.read()
    try:
        decrypted = fernet.decrypt(encrypted)
        with open(file_path, 'wb') as decrypted_file:
            decrypted_file.write(decrypted)
        return True
    except Exception as e:
        return False

# Checks to see if path is a directory or file, then processes for encryption/decryption
def process_files(action, path, key):
    if os.path.isfile(path):
        if action == 'encrypt':
            encrypt_file(path, key)
        elif action == 'decrypt':
            if not try_decrypt_file(path, key):
                print("Incorrect decryption key.")
                return False
    elif os.path.isdir(path):
        for root, dirs, files in os.walk(path):
            for file in files:
                file_path = os.path.join(root, file)
                if action == 'encrypt':
                    encrypt_file(file_path, key)
                elif action == 'decrypt':
                    if not try_decrypt_file(file_path, key):
                        print(f"Incorrect decryption key for {file_path}.")
                        return False
    else:
        print("Invalid file or directory path.")
        return False
    return True

# Check if the key can be base64 decoded and has the correct length IF REMOVED CODE WILL CRASH
def is_valid_key(key):
    try:
        decoded_key = base64.urlsafe_b64decode(key)
        return len(decoded_key) == 32
    except Exception:
        return False

# Create an infinite loop so the script keeps running until user prompts to exit
# allow 3 valid attempts to choose, then start prompting user to exit script
while True:
    action = input("Would you like to encrypt or decrypt files? ").strip().lower()
    attempts = 3
    while action not in ['encrypt', 'decrypt'] and attempts > 0:
        attempts -= 1
        print("Invalid input. Please type 'encrypt' or 'decrypt'.")
        action = input("Would you like to encrypt or decrypt files? ").strip().lower()
        # option to exit after too many invalid attempts
        if attempts == 0:
            exit_choice = input("Too many invalid attempts. Do you want to exit? (yes/no): ").strip().lower()
            if exit_choice == 'yes':
                print("Exiting script.")
                sys.exit()
            else:
                attempts = 1

# Starts encrpytion process
    if action == 'encrypt':
        key = generate_key()
        print("Encryption Key: ", key.decode())
        input("Press Enter after you have copied the encryption key.")
    
# asks user for decrytion key to start decryption process
    else:
        key = input("Enter the encryption key (leave blank to exit): ").strip()
        if key == "":
            print("No key provided. Exiting script.")
            sys.exit()

# makes sure the key is valid, else script crashes
        while not is_valid_key(key):
            print("Invalid key.")
            key = input("Enter the correct encryption key (leave blank to exit): ").strip()
            if key == "":
                print("No key provided. Exiting script.")
                sys.exit()
        key = key.encode()

# Handles the file or directory paths, then processes files for encryption/decryption
    while True:
        path = input("Enter the path of the file or directory: ").strip()
        if process_files(action, path, key):
            print(f"{action.capitalize()}ion completed successfully.")
            break
        else:
            retry_choice = input("File path does not exist. Do you want to try again or exit? (try again/exit): ").strip().lower() # if file path was invlaid, retries
            if retry_choice == 'exit':
                print("Exiting script.")
                sys.exit()
            elif retry_choice == 'try again':
                if action == 'decrypt':
                    key = input("Enter the encryption key (leave blank & press enter to exit): ").strip() # prompts user for decryption key and makes sure key is correct.
                    if key == "":
                        print("No key provided. Exiting script.")
                        sys.exit()
                    while not is_valid_key(key):
                        print("Invalid key.")
                        key = input("Enter the correct encryption key (leave blank & press enter to exit): ").strip()
                        if key == "":
                            print("No key provided. Exiting script.")
                            sys.exit()
                    key = key.encode()
            else:
                print("Invalid input. Please type 'try again' or 'exit'.")

# ask user if they want process another file or exit the script
    next_action = input("Do you want to encrypt/decrypt another file or directory? (yes/no): ").strip().lower()
    while next_action not in ['yes', 'no']:
        print("Invalid input. Please type 'yes' or 'no'.")
        next_action = input("Do you want to encrypt/decrypt another file or directory? (yes/no): ").strip().lower()
    if next_action == 'no':
        print("Exiting script.")
        break
