# To use this Makefile, get a copy of my SF Release Tools
# git clone git://git.code.sf.net/p/sfreleasetools/code sfreleasetools
# And point the environment variable RELEASETOOLS to the checkout

ifeq (,${RELEASETOOLS})
    RELEASETOOLS=../releasetools
endif
README:=README.rst
PKG=ooopy
PY=__init__.py OOoPy.py Transformer.py Transforms.py
SRC=Makefile MANIFEST.in setup.py $(README) README.html \
    $(PY:%.py=$(PKG)/%.py) testfiles/* bin/*

VERSIONPY=ooopy/Version.py
VERSION=$(VERSIONPY)
LASTRELEASE:=$(shell $(RELEASETOOLS)/lastrelease -n)

USERNAME=schlatterbeck
PROJECT=ooopy
PACKAGE=${PKG}
CHANGES=changes
NOTES=notes

all: $(VERSION)

$(VERSION): $(SRC)

test: $(VERSION)
	$(PYTHON) run_doctest.py ooopy/OOoPy.py
	$(PYTHON) run_doctest.py ooopy/Transforms.py
	$(PYTHON) run_doctest.py ooopy/Transformer.py

clean:
	rm -f $(PKG)/Version.pyc $(PKG)/testout.sxw $(PKG)/testout2.sxw
	rm -f testout.sxw testout.odt testout2.sxw testout2.odt \
	    testout3.sxw testout3.odt out.html out2.odt         \
	    out.sxw carta-out.stw carta-out.odt xyzzy.odt
	rm -rf $(PKG)/__pycache__ __pycache__
	rm -f ooopy/Version.py ooopy/Version.py{c,o} 
	rm -f $(PKG)/Version.py
	rm -rf ${CLEAN}

release: upload upload_homepage announce_pypi announce

include $(RELEASETOOLS)/Makefile-pyrelease
