default: help

#help: @ Shows help topics
help:
	@grep -E '[a-zA-Z\.\-]+:.*?@ .*$$' $(MAKEFILE_LIST)| tr -d '#'  | awk 'BEGIN {FS = ":.*?@ "}; {printf "\033[32m%-30s\033[0m%s\n", $$1, $$2}'

#install: @ Clean and install deps
install: clean
	pip3 install -r requirements.txt

#test: @ Run all tests
test:
	python3 -m unittest discover -p "test_*.py"

#clean: @ Remove build, dist and cache files
clean: clean-build clean-pyc clean-test

#clean-build: @ Remove build files
clean-build:
	rm -fr build/
	rm -fr dist/
	rm -fr .eggs/
	find . -name '*.egg-info' -exec rm -fr {} +
	find . -name '*.egg' -exec rm -f {} +

#clean-pyc: @ Remove cache files
clean-pyc:
	find . -name '*.pyc' -exec rm -f {} +
	find . -name '*.pyo' -exec rm -f {} +
	find . -name '*~' -exec rm -f {} +
	find . -name '__pycache__' -exec rm -fr {} +

#clean-test: @ Remove tests generated files
clean-test:
	rm -fr .tox/
	rm -f .coverage
	rm -fr htmlcov/

#release: @ Build and Release new version
release: clean dist
	python3 setup.py sdist
	python3 setup.py bdist_wheel
	python3 -m twine upload dist/*

#dist: @ Build dist version
dist: clean
	python3 setup.py sdist
	python3 setup.py bdist_wheel
	ls -l dist

.PHONY: install test clean release dist
