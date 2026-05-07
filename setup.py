from setuptools import setup, find_packages

setup(
    name='email-automation',
    version='1.0.0',
    description='Email automation system for processing Excel attachments to SQL Server',
    author='Your Name',
    packages=find_packages(),
    install_requires=[
        'pandas>=2.0.0',
        'openpyxl>=3.1.0',
        'pyodbc>=4.0.39',
        'python-dotenv>=1.0.0',
        'schedule>=1.2.0',
    ],
    python_requires='>=3.9',
    entry_points={
        'console_scripts': [
            'email-automation=main:main',
        ],
    },
)
