dist:
	python setup.py sdist bdist_wheel

clean:
	-rm -rf docxbuilder/docx/style.docx build/ dist/ *.egg-info

upload: clean dist
	python -m twine upload --repository pypi dist/*

.PHONY: dist clean upload
