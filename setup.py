from setuptools import setup, find_packages

setup(
    name='frontacc_conv',
    version='0.1.1',
    packages=find_packages(),
    install_requires=[
        'pandas',
        'xlrd >= 2.0.1'
    ],
    py_modules=["frontacc_conv"],  # Refers to the frontacc_conv.py file
    entry_points={
        'console_scripts': [
            'frontacc_conv=frontacc_conv:main',  # Entry point for command line
        ],
    },
    description='Convert GL Account Transactions to QIF format',
    author='Your Name',
    author_email='your.email@example.com',
    url='https://yourmodulehomepage.com',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)
