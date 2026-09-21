"""Exclusive output reservation and atomic writes shared by CLI and converters."""
import os
import tempfile


def reserve_output(path_factory):
    """Claim a name before conversion; competing processes must pick another."""
    while True:
        path = path_factory()
        try:
            with open(path, 'xb'):
                pass
            return path
        except FileExistsError:
            continue


def atomic_write(path, content):
    """Replace an owned reservation only after the complete content is on disk."""
    fd, temporary = tempfile.mkstemp(prefix='.bdc-', dir=os.path.dirname(path))
    try:
        with os.fdopen(fd, 'w', encoding='utf-8', newline='') as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
