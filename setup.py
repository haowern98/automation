"""
Setup script for SharePoint Automation
"""
from setuptools import setup, find_packages

setup(
    name="sharepoint_automation",
    version="1.0.0",
    description="SharePoint Automation tool for GSN vs ER reports",
    author="Your Name",
    packages=find_packages(),
install_requires=[
    "pandas>=1.3.0",
    "openpyxl>=3.0.7",
    "xlwings>=0.25.0",
    "pywin32>=301",
    "ldap3>=2.9.1",
    "PyQt5>=5.15.4",
    "pyad>=0.6.0",
    "winshell>=0.6.0"  # Added this line
],
    entry_points={
        'console_scripts': [
            'sharepoint-automation=main:main',
        ],
    },
    python_requires='>=3.7',
)