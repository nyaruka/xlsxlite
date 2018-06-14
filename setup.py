# Always prefer setuptools over distutils
from setuptools import setup, find_packages
from os import path

here = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(here, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()


def _is_requirement(line):
    """Returns whether the line is a valid package requirement."""
    line = line.strip()
    return line and not line.startswith("#")


def _read_requirements(filename):
    """Parses a file for pip installation requirements."""
    with open(filename) as requirements_file:
        contents = requirements_file.read()
    return [line.strip() for line in contents.splitlines() if _is_requirement(line)]


setup(
    name='xlsxlite',
    version=__import__('xlsxlite').__version__,
    description='Lightweight XLSX writer with emphasis on minimizing memory usage.',
    long_description=long_description,
    long_description_content_type='text/markdown',

    classifiers=[
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Programming Language :: Python'
    ],
    keywords='excel xlxs',
    url='http://github.com/nyaruka/xlxslite',
    license='MIT',

    author='Nyaruka',
    author_email='code@nyaruka.com',

    packages=find_packages(),
    install_requires=_read_requirements("requirements/base.txt"),
    tests_require=_read_requirements("requirements/tests.txt"),

    project_urls={
        'Bug Reports': 'https://github.com/nyaruka/xlsxlite/issues',
        'Source': 'https://github.com/nyaruka/xlsxlite/',
    },
)
