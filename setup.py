import setuptools
from pathlib import Path

this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text()


setuptools.setup(
    name="xxlSearch",
    version="0.1.0",
    description="Search for words in Excel files",
    url="https://github.com/willros/xxlSearch",
    author="William Rosenbaum",
    author_email="william.rosenbaum@gmail.com",
    license="MIT",
    packages=setuptools.find_packages(),
    long_description=long_description,
    long_description_content_type="text/markdown",
    python_requires=">=3.10",
    install_requires=[
        "pandas==2.2.1",
        "openpyxl==3.1.2",
    ],
    entry_points={"console_scripts": ["xxlSearch=xxlSearch.main:main"]},
)