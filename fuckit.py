from hashlib import sha256
h = sha256()
h.update(b'abcd1234')
hash = h.hexdigest()
print(hash)