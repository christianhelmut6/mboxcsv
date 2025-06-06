import modal
import sys

print(f"Python version: {sys.version}")
print(f"Modal library location: {modal.__file__}")
print(f"Modal library version: {modal.__version__}")

try:
    print(f"modal.Mount available: {hasattr(modal, 'Mount')}")
    if hasattr(modal, 'Mount'):
        print(f"modal.Mount type: {type(modal.Mount)}")
except Exception as e:
    print(f"Error accessing modal.Mount: {e}")
