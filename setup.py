from setuptools import setup, find_packages

setup(
    name='outlook_excel_mailer',
    version='1.0.0',
    author='Your Name',
    author_email='your.email@example.com',
    description='Send Outlook emails based on Excel milestone hyperlinks',
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        'openpyxl==3.1.2',
        'pywin32==306',
        'python-dotenv==1.0.1'
    ],
    entry_points={
        'console_scripts': [
            'send-mail=main:main'
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: Microsoft :: Windows",
    ],
)
