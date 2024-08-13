from setuptools import setup, find_packages

setup(
    name="pdf_to_excel",
    version="0.1.0",
    packages=find_packages(),
    include_package_data=True,
    package_data={
        "package": ["base.xls", "car-image.png"],
    },
    entry_points={
        "console_scripts": [
            "pdf_to_excel=package.main:main",
        ],
    },
    install_requires=[
        "pdfplumber",
        "pandas",
        "xlwt",
        "xlrd",
        "xlutils",
        "xlwings",
        "xlsxwriter",
        "openpyxl",
    ],
)
