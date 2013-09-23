PKG=ooopy
PY=__init__.py OOoPy.py Transformer.py Transforms.py
SRC=Makefile MANIFEST.in setup.py README README.html \
    $(PY:%.py=$(PKG)/%.py) testfiles/* bin/*

VERSIONPY=ooopy/Version.py
VERSION=$(VERSIONPY)
LASTRELEASE:=$(shell ../svntools/lastrelease -n)

USERNAME=schlatterbeck
PROJECT=ooopy
PACKAGE=${PKG}
CHANGES=changes
NOTES=notes

all: $(VERSION)

$(VERSION): $(SRC)

dist: all
	python setup.py sdist --formats=gztar,zip

test: $(VERSION)
	python run_doctest.py ooopy/OOoPy.py
	python run_doctest.py ooopy/Transforms.py
	python run_doctest.py ooopy/Transformer.py

clean:
	rm -f $(PKG)/Version.pyc $(PKG)/testout.sxw \
	    $(PKG)/testout2.sxw ${CHANGES} ${NOTES}
	rm -rf dist build
	rm -f testout.sxw testout.odt testout2.sxw testout2.odt \
	    testout3.sxw testout3.odt out.html out2.odt         \
	    out.sxw carta-out.stw carta-out.odt xyzzy.odt       \
            upload upload_homepage announce_pypi announce

clobber: clean
	rm -f $(PKG)/Version.py MANIFEST README.html default.css

release: upload upload_homepage announce_pypi announce

include ../make/Makefile-sf
