PKG=ooopy
PY=__init__.py OOoPy.py Transformer.py Transforms.py
SRC=Makefile MANIFEST.in setup.py README README.html default.css \
    $(PY:%.py=$(PKG)/%.py) test.sxw test.odt

VERSION=ooopy/Version.py
LASTRELASE:=$(shell ../svntools/lastrelease -n)

USERNAME=schlatterbeck
PROJECT=ooopy
PACKAGE=${PKG}
CHANGES=changes
NOTES=notes
HOSTNAME=shell.sourceforge.net
PROJECTDIR=/home/groups/o/oo/ooopy/htdocs

all: $(VERSION)

$(VERSION): $(SRC)

dist: all
	python setup.py sdist --formats=gztar,zip

README.html: README default.css
	rst2html --stylesheet=default.css $< > $@

default.css: ../../content/html/stylesheets/default.css
	cp ../../content/html/stylesheets/default.css .

%.py: %.v $(SRC)
	sed -e 's/RELEASE/$(LASTRELASE)/' $< > $@

upload_homepage: all
	scp README.html $(USERNAME)@$(HOSTNAME):$(PROJECTDIR)/index.html
	scp default.css $(USERNAME)@$(HOSTNAME):$(PROJECTDIR)

clean:
	rm -f MANIFEST README.html default.css \
	    $(PKG)/Version.py $(PKG)/Version.pyc $(PKG)/testout.sxw \
	    $(PKG)/testout2.sxw ${CHANGES} ${NOTES}
	rm -rf dist build

include ../make/Makefile-sf
