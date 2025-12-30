import hashlib
import os

# Password to hash
password = "mrsi123"

# Generate a random salt (a random value added to the password before hashing)
# A strong salt should be at least 16 bytes (16 * 2 = 32 hex chars)
salt = os.urandom(16)
print(f"Generated Salt (Hex): {salt.hex()}")

# Hash the password + salt
hashed_password = hashlib.sha256(salt + password.encode('utf-8')).hexdigest()
print(f"Hashed Password (SHA-256): {hashed_password}")

# IMPORTANT: You must use these exact strings in your application code.
# For example:
# SALT = b'YOUR_GENERATED_SALT_BYTES' # Note the 'b' prefix for bytes
# HASH = 'YOUR_GENERATED_HASH_STRING'