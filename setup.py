from setuptools import setup
from Cython.Build import cythonize

setup(
    ext_modules=cythonize(
        ['Priva_Data_Transfer.py'],  # Python code file with primes() function
        annotate=True),                 # enables generation of the html annotation file
)