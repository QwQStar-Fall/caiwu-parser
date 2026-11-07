from setuptools import setup, find_packages

with open("README.md", encoding="utf-8") as f:
    long_description = f.read()

setup(
    name="caiwu-parser",
    version="0.1.0",
    description="轻量、高鲁棒性的 A 股上市公司财报三大报表解析器（支持原生 PDF / Excel / Word）",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Star-Fall",
    url="https://github.com/QwQStar-Fall/caiwu-parser",
    license="MIT",
    packages=find_packages(include=["caiwu_parser*"]),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Topic :: Office/Business :: Financial :: Accounting",
    ],
    python_requires=">=3.8",
    install_requires=[
        "pandas>=1.3.0",
        "pdfplumber>=0.7.0",
        "pypdfium2>=4.0.0",
        "openpyxl>=3.0.0",
        "python-docx>=0.8.11",
    ],
    entry_points={
        "console_scripts": [
            "caiwu-parser=caiwu_parser.__main__:main",
        ],
    },
)
