LASTRELASE:=$(shell if x=`lastrelease -d` ;then echo $$x ;else echo 'NO_TAG' ;fi)
OOOPY=__init__.py OOoPy.py Transformer.py Transforms.py
VERSION=ooopy/Version.py
SRC=default.css Makefile MANIFEST.in README setup.py test.sxw \
    $(OOOPY:%.py=ooopy/%.py) README.html


all: $(VERSION)

$(VERSION): $(SRC)

README.html: README
	rst2html $< > $@

dist: all
	python setup.py sdist

%.py: %.v
	sed -e 's/RELEASE/$(LASTRELASE)/' $< > $@
