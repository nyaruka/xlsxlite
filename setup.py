from setuptools import setup, find_packages

try:
    from pypandoc import convert
    read_md = lambda f: convert(f, 'rst')  # noqa
except ImportError:
    print("warning: pypandoc module not found, could not convert Markdown to RST")
    read_md = lambda f: open(f, 'r').read()  # noqa


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
    long_description=read_md('README.md'),

    classifiers=[
        'Environment :: Web Environment',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Programming Language :: Python'
    ],
    keywords='excel xlxs',
    url='http://github.com/nyaruka/xlxslite',
    license='MIT',

    maintainer='Nyaruka',
    maintainer_email='code@nyaruka.com',

    packages=find_packages(),
    install_requires=_read_requirements("requirements/base.txt"),
    tests_require=_read_requirements("requirements/tests.txt")
)
