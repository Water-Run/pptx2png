# setup.py of pptx2png

import os
from setuptools import setup, find_packages

# Read long description safely
long_desc = ""
if os.path.exists("README.md"):
    with open("README.md", "r", encoding="utf-8") as f:
        long_desc = f.read()

setup(
    name="pptx2png",
    version="2025.1",
    author="WaterRun",
    author_email="2263633954@qq.com",
    description="A one-click Python library for converting .pptx files into .png images.",
    long_description=long_desc,
    long_description_content_type="text/markdown",
    url="https://github.com/Water-Run/pptx2png",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Topic :: Multimedia :: Graphics :: Presentation",
    ],
    install_requires=[
        "pywin32",
    ],
    python_requires='>=3.5',
)