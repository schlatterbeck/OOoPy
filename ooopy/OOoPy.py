#!/usr/bin/env python
from zipfile                 import ZipFile, ZIP_DEFLATED
from StringIO                import StringIO
from elementtree.ElementTree import ElementTree, fromstring
from tempfile                import mkstemp
import os

def copy (srcfile, dstfile = None, mode = 'a') :
    """
        Factory function to produce a OOoPy object from a *copy* of the
        given src zip file. This opens the original in read-mode, copies
        everything over to the new and then opens the new one in write
        mode. If no destination is given, we create a temporary file.
    """
    if not dstfile :
        f, dstfile = mkstemp ('.sxw', 'ooopy')
        os.close (f)
    src = ZipFile (srcfile, mode = "r")
    dst = OOoPy   (dstfile, mode = "w")
    for f in src.infolist () :
        dst.zip.writestr (f, src.read (f.filename))
    src.close ()
    dst.close ()
    return OOoPy   (dstfile, mode = mode)
# end def copy

class OOoElementTree (object) :
    """
        An ElementTree for OOo document XML members. Behaves like the
        orginal ElementTree (in fact it delegates almost everything to a
        real instance of ElementTree) expect for the write method, that
        writes itself back to the OOo XML file in the OOo zip archive it
        came from.
    """
    def __init__ (self, ooopy, zname, root) :
        self.ooopy = ooopy
        self.zname = zname
        self.tree  = ElementTree (root)
    # end def __init__

    def write (self) :
        self.ooopy.write (self.zname, self.tree)
    # end def write

    def __getattr__ (self, name) :
        """
            Delegate everything to our ElementTree attribute.
        """
        if not name.startswith ('__') :
            return getattr (self.tree, name)
        raise AttributeError, name
    # end def __getattr__

# end class OOoElementTree

class OOoPy (object) :
    """
        Wrapper for OpenOffice.org zip files (all OOo documents are
        really zip files internally).

        from OOoPy import OOoPy
        >>> o = OOoPy ('test.sxw', mode = 'a')
        >>> e = o.read ('content.xml')
        >>> e.write ()
    """
    def __init__ (self, filename, mode = 'r') :
        """
            Open an OOo document with the given mode. Default is "r" for
            read-only access.
        """
        self.filename = filename
        self.zip      = ZipFile (filename, mode, ZIP_DEFLATED)
    # end def __init__

    def read (self, zname) :
        """
            return an OOoElementTree object for the given OOo document
            archive member name. Currently an OOo document contains the
            following XML files::

             * content.xml: the text of the OOo document
             * styles.xml: style definitions
             * meta.xml: meta-information (author, last changed, ...)
             * settings.xml: settings in OOo
             * META-INF/manifest.xml: contents of the archive

            There is an additional file "mimetype" that always contains
            the string "application/vnd.sun.xml.writer".
        """
        return OOoElementTree (self, zname, fromstring (self.zip.read (zname)))
    # end def tree

    def write (self, zname, etree) :
        str = StringIO ()
        etree.write (str)
        self.zip.writestr (zname, str.getvalue ())
    # end def

    def close (self) :
        """
            Close the zip file. According to documentation of zipfile in
            the standard python lib, this has to be done to be sure
            everything is written.
        """
        self.zip.close ()
    # end def close
# end class OOoPy
