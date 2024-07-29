from hashlib import sha256
h = sha256()
h.update(b'grand')
hash = h.hexdigest()
print(hash)