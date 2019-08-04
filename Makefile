dist:
	python setup.py sdist
	python setup.py bdist_wheel

clean:
	-rm -rf docxbuilder/docx/style.docx build/ dist/ *.egg-info

upload: clean dist
	python -m twine upload --repository pypi dist/*

test: clean dist
	python -m twine upload --repository testpypi dist/*

.PHONY: dist clean upload test
