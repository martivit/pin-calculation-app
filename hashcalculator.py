from hashlib import sha256
h = sha256()
h.update(b'martina')
hash = h.hexdigest()
print(hash)