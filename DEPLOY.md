```
pip install wheel twine
rm -R dist/
python setup.py sdist
python setup.py bdist_wheel
twine upload -u nicpottier dist/*
```
