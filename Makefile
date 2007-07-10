PKG=ooopy
PY=__init__.py OOoPy.py Transformer.py Transforms.py
SRC=Makefile MANIFEST.in setup.py README README.html default.css \
    $(PY:%.py=$(PKG)/%.py) test.sxw test.odt

VERSION=ooopy/Version.py
LASTRELASE:=$(shell ../svntools/lastrelease)

USERNAME=schlatterbeck
HOSTNAME=shell.sourceforge.net
PROJECTDIR=/home/groups/o/oo/ooopy/htdocs

all: $(VERSION)

$(VERSION): $(SRC)

dist: all
	python setup.py sdist --formats=gztar,zip

README.html: README
	rst2html $< > $@

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
	    $(PKG)/testout2.sxw
	rm -rf dist build
