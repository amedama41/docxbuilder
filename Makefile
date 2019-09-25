dist:
	python setup.py sdist
	python setup.py bdist_wheel --universal

clean:
	-rm -rf docxbuilder/docx/style.docx build/ dist/ *.egg-info

update_style_file:
	./create_style_file.py
	(cd style_file; make docx)
	mv style_file/build/docx/style.docx docxbuilder/docx/style.docx
	(cd style_file; make docx)
	unzip style_file/build/docx/style.docx -d style_file/docx

upload: clean dist
	python -m twine upload --repository pypi dist/*

test: clean dist
	python -m twine upload --repository testpypi dist/*

.PHONY: dist clean upload test
