import base64

def decrypt(password):
    password = str(base64.b64decode(password))
    return password.replace('b', "").replace("'", "")