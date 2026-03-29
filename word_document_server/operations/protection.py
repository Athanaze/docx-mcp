"""
Document encryption/decryption using msoffcrypto-tool.

Provides real password protection (not simulated sidecar files).
"""
import os
import io
import tempfile

from word_document_server.document import raw_docx_tool


@raw_docx_tool()
def encrypt_document(filename, password):
    """Encrypt a Word document with a password using msoffcrypto."""
    import msoffcrypto

    with open(filename, 'rb') as f:
        file_content = f.read()

    encrypted = io.BytesIO()
    file_obj = io.BytesIO(file_content)
    ms_file = msoffcrypto.OfficeFile(file_obj)
    ms_file.load_key(password=password)
    ms_file.encrypt(password, encrypted)

    with open(filename, 'wb') as f:
        f.write(encrypted.getvalue())

    return f"Document {filename} encrypted successfully"


@raw_docx_tool()
def decrypt_document(filename, password):
    """Decrypt a password-protected Word document."""
    import msoffcrypto

    with open(filename, 'rb') as f:
        file_obj = io.BytesIO(f.read())

    try:
        ms_file = msoffcrypto.OfficeFile(file_obj)
        ms_file.load_key(password=password)
    except Exception:
        return f"Failed to decrypt: incorrect password or file is not encrypted"

    decrypted = io.BytesIO()
    ms_file.decrypt(decrypted)

    with open(filename, 'wb') as f:
        f.write(decrypted.getvalue())

    return f"Document {filename} decrypted successfully"
