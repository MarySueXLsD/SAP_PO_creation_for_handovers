from cryptography.fernet import Fernet
import base64
import configparser

key = Fernet.generate_key()  # Save this key somewhere safe
print(key)
cipher_suite = Fernet(key)

# Encrypt the data
#username = b"TESTROBOT"
#password = b"TESTROBOT"
username = b"PRODROBOT"
password = b"PRODROBOT"
encrypted_username = cipher_suite.encrypt(username)
encrypted_password = cipher_suite.encrypt(password)

# Convert the encrypted bytes to a base64 encoded string
encrypted_username_str = base64.urlsafe_b64encode(encrypted_username).decode('utf-8')
encrypted_password_str = base64.urlsafe_b64encode(encrypted_password).decode('utf-8')

# Store the encrypted strings in the config file
config = configparser.ConfigParser()
config['SAP_login'] = {'username': encrypted_username_str, 'password': encrypted_password_str}

with open('static/config.ini', 'w') as configfile:
    config.write(configfile)
