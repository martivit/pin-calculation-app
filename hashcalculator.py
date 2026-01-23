import hmac

# Example password and secret key
password = 'pinsession'
SECRET_KEY = b'supersecretkey'

# Create the HMAC hash using the secret key and the password
hash_value = hmac.new(SECRET_KEY, password.encode('utf-8'), 'sha256').hexdigest()

# Print the resulting hash
print(hash_value)
