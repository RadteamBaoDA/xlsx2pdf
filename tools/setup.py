"""
Setup configuration for Office to PDF Converter package
"""
from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="office-to-pdf-converter",
    version="2.0.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="Convert Microsoft Office files (Excel, Word, PowerPoint) to PDF with advanced formatting options",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/xlsx2pdf",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Topic :: Office/Business",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: Microsoft :: Windows",
    ],
    python_requires=">=3.7",
    install_requires=[
        "pywin32>=300",
        "rich>=10.0.0",
        "pyyaml>=5.4.0",
        "psutil>=5.8.0",
        "pypdf>=3.0.0",
    ],
    extras_require={
        "dev": ["pytest>=6.0", "black", "flake8"],
        "language": ["langdetect>=1.0.9"],
    },
    entry_points={
        "console_scripts": [
            "office2pdf=main:main",
        ],
    },
)
