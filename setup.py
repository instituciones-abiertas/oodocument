import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="oodocument",
    version="0.0.5",
    author="José Luis Di Biase",
    author_email="josx@camba.coop",
    description="Basic tasks to convert and replace files using uno api",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Cambalab/oodocument",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        "Operating System :: OS Independent",
        "Topic :: Office/Business :: Office Suites",
    ],
    python_requires='>=3.6',
)
