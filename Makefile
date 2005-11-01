PKG=ooopy
PY=__init__.py OOoPy.py Transformer.py Transforms.py
SRC=Makefile MANIFEST.in setup.py README README.html default.css \
    $(PY:%.py=$(PKG)/%.py) test.sxw

VERSION=ooopy/Version.py
LASTRELASE:=$(shell if x=`lastrelease -d` ;then echo $$x ;else echo 'NO_TAG' ;fi)

USERNAME=schlatterbeck
HOSTNAME=shell.sourceforge.net
PROJECTDIR=/home/groups/o/oo/ooopy/htdocs

all: $(VERSION)

$(VERSION): $(SRC)

dist: all
	python setup.py sdist --formats=gztar,zip

README.html: README
	rst2html $< > $@

default.css: ../../html/stylesheets/default.css
	cp ../../html/stylesheets/default.css .

%.py: %.v
	sed -e 's/RELEASE/$(LASTRELASE)/' $< > $@

upload_homepage: all
	scp README.html $(USERNAME)@$(HOSTNAME):$(PROJECTDIR)/index.html
	scp default.css $(USERNAME)@$(HOSTNAME):$(PROJECTDIR)

clean:
	rm -f MANIFEST README.html default.css \
	    $(PKG)/Version.py $(PKG)/Version.pyc $(PKG)/testout.sxw \
	    $(PKG)/testout2.sxw
	rm -rf dist build
