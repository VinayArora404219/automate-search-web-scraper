from setuptools import setup

setup(
    name='automate-search',
    version='1.0',
    packages=[''],
    url='',
    license='',
    author='vinay',
    author_email='',
    description='',
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
    ],
install_requires=[
        "html2text","undetected-chromedriver","requests","python-docx",
    ],
    entry_points={"console_scripts": ["__main__:main"]}
)
