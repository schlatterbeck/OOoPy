PKG=ooopy
PY=__init__.py OOoPy.py Transformer.py Transforms.py
SRC=Makefile MANIFEST.in setup.py README README.html \
    $(PY:%.py=$(PKG)/%.py) test.sxw test.odt

VERSION=ooopy/Version.py
LASTRELEASE:=$(shell ../svntools/lastrelease -n)

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

test: $(VERSION)
	python run_doctest.py ooopy/OOoPy.py
	python run_doctest.py ooopy/Transforms.py
	python run_doctest.py ooopy/Transformer.py

clean:
	rm -f MANIFEST README.html default.css \
	    $(PKG)/Version.py $(PKG)/Version.pyc $(PKG)/testout.sxw \
	    $(PKG)/testout2.sxw ${CHANGES} ${NOTES}
	rm -rf dist build
	rm -f testout.sxw testout.odt testout2.sxw testout2.odt \
	    testout3.sxw testout3.odt out.html                  \
	    out.sxw carta-out.stw carta-out.odt

release: upload upload_homepage announce_pypi announce

include ../make/Makefile-sf
