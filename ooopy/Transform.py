#!/usr/bin/env python
import OOoPy
import time
import re
from elementtree.ElementTree import dump, SubElement, Element

files = ['content.xml', 'styles.xml', 'meta.xml', 'settings.xml']
tags = \
{ 'chart'  : "http://openoffice.org/2000/chart"
, 'config' : "http://openoffice.org/2001/config"
, 'dc'     : "http://purl.org/dc/elements/1.1/"
, 'dr3d'   : "http://openoffice.org/2000/dr3d"
, 'draw'   : "http://openoffice.org/2000/drawing"
, 'fo'     : "http://www.w3.org/1999/XSL/Format"
, 'form'   : "http://openoffice.org/2000/form"
, 'math'   : "http://www.w3.org/1998/Math/MathML"
, 'meta'   : "http://openoffice.org/2000/meta"
, 'number' : "http://openoffice.org/2000/datastyle"
, 'office' : "http://openoffice.org/2000/office"
, 'script' : "http://openoffice.org/2000/script"
, 'style'  : "http://openoffice.org/2000/style"
, 'svg'    : "http://www.w3.org/2000/svg"
, 'table'  : "http://openoffice.org/2000/table"
, 'text'   : "http://openoffice.org/2000/text"
, 'xlink'  : "http://www.w3.org/1999/xlink"
}

def OOo_Tag (tag, name) :
    """Return combined XML tag"""
    return "{%s}%s" % (tags [tag], name)
# end def OOo_Tag

class Transform (object) :
    """Base class for individual transforms on OOo files. An individual
       transform needs a filename variable for specifying the OOo file
       the transform should be applied to and an optional prio.
       Individual transforms are applied according to their prio
       setting, higher prio means later application of a transform.

       The filename variable must specify one of the XML files which are
       part of the OOo document (see files variable above). As
       the names imply, content.xml contains the contents of the
       document (text and ad-hoc style definitions), styles.xml contains
       the style definitions, meta.xml contains meta information like
       author, editing time, etc. and settings.xml is used to store
       OOo's settings (menu Tools->Configure).
    """
    prio = 100
    def __init__ (self, prio = None) :
        if prio :
            self.prio    = prio
        self.transformer = None
    # end def __init__

    def register (self, transformer) :
        self.transformer = transformer
        transformer.register (self)
    # end def register

    def _classname_prefix (self) :
        """For fulfilling the naming convention of the transformer
           dictionary (every entry in this dictionary should be prefixed
           with the class name of the transform) we have this
           convenience method.
        """
        return self.__class__.__name__ + ':'
    # end def _classname_prefix

    classname_prefix = property (_classname_prefix)

    def apply (self, root, transformer) :
        raise NotImplementedError, 'derived transforms must implement "apply"'
    # end def apply

# end class Transform

class Transformer (object) :
    """Class for applying a set of transforms to a given ooopy object.
       The transforms are applied to the specified file in priority
       order. When applying transforms we have a mechanism for
       communication of transforms. We give the transformer to the
       individual transforms as a parameter. The transforms may use the
       transformer like a dictionary for storing values and retrieving
       values left by previous transforms.
       As a naming convention each transform should use its class name
       as a prefix for storing values in the dictionary.
    """
    def __init__ (self, *ts) :
        self.transforms = {}
        for f in files :
            self.transforms [f] = {}
        for t in ts :
            self.insert (t)
        self.dictionary = {}
        self.has_key = self.dictionary.has_key
    # end def __init__

    def insert (self, t) :
        """Insert a new transform"""
        if not self.transforms [t.filename].has_key (t.prio) :
            self.transforms [t.filename][t.prio] = []
        self.transforms [t.filename][t.prio].append (t)
    # end def append

    def transform (self, ooopy) :
        """Apply all the transforms in priority order"""
        for f in self.transforms.keys () :
            if self.transforms [f] :
                e = ooopy.read (f)
                root = e.getroot ()
                prios = self.transforms [f].keys ()
                prios.sort ()
                for prio in prios :
                    for t in self.transforms [f][prio] :
                        t.apply (root, self)
                e.write ()
    # end def transform

    def __getitem__ (self, key) :
        return self.dictionary [key]
    # end def __getitem__

    def __setitem__ (self, key, value) :
        self.dictionary [key] = value
    # end def __setitem__
# end class Transformer

#
# meta.xml transforms
#

class Editinfo_Transform (Transform) :
    """This is an example of modifying OOo meta info (edit information,
       author, etc). We set some of the items (program that generated
       the OOo file, modification time, number of edit cyles and overall
       edit duration).  It's easy to subclass this transform and replace
       the "replace" variable (pun intended) in the derived class.
    """
    filename = 'meta.xml'
    prio     = 20
    replace = \
    { OOo_Tag ('meta', 'generator')        : 'OOoPy field replacement'
    , OOo_Tag ('dc',   'date')             : time.strftime ('%Y-%m-%dT%H:%M:%S')
    , OOo_Tag ('meta', 'editing-cycles')   : '0'
    , OOo_Tag ('meta', 'editing-duration') : 'PT0M0S'
    }

    def apply (self, root, transformer) :
        for node in root.findall (OOo_Tag ('office', 'meta') + '/*') :
            if self.replace.has_key (node.tag) :
                node.text = self.replace [node.tag]
    # end def apply
# end class Editinfo_Transform

#
# settings.xml transforms
#

class Autoupdate_Transform (Transform) :
    """This is an example of modifying OOo settings. We set some of the
       AutoUpdate configuration items in OOo to true. We also specify
       that links should be updated when reading.

       This was originally intended to make OOo correctly display fields
       if they were changed with the Field_Replace_Transform below
       (similar to pressing F9 after loading the generated document in
       OOo). In particular I usually make spaces depend on field
       contents so that I don't have spurious spaces if a field is
       empty. Now it would be nice if OOo displayed the spaces correctly
       after loading a document (It does update the fields before
       printing, so this is only a cosmetic problem :-). This apparently
       does not work. If anybody knows how to achieve this, please let
       me know: mailto:rsc@runtux.com
    """
    filename = 'settings.xml'
    prio     = 20

    def apply (self, root, transformer) :
        config = None
        for config in root.findall \
            ( OOo_Tag ('office', 'settings')
            + '/'
            + OOo_Tag ('config', 'config-item-set')
            ) :
            name = config.get (OOo_Tag ('config', 'name'))
            if name == 'configuration-settings' :
                break
        for node in config.findall (OOo_Tag ('config', 'config-item')) :
            name = node.get (OOo_Tag ('config', 'name'))
            if name == 'LinkUpdateMode' :  # update when reading
                node.text = '2'
            # update fields when reading
            if name == 'FieldAutoUpdate' or name == 'ChartAutoUpdate' :
                node.text = 'true'
    # end def apply
# end class Autoupdate_Transform

#
# content.xml transforms
#

class Field_Replace_Transform (Transform) :
    """ Takes a dict of replacement key-value pairs. The key is the name
        of a variable in OOo. Additional replacement key-value pairs may
        be specified in **kw. Alternatively a callback mechanism for
        variable name lookups is provided. The callback function is
        given the name of a variable in OOo and is expected to return
        the replacement value or None if the variable value should not
        be replaced.
    """
    filename = 'content.xml'
    prio     = 100

    def __init__ (self, prio = None, replace = None, callback = None, **kw) :
        Transform (prio)
        self.replace  = replace or {}
        self.callback = callback
        for k in kw.keys () :
            self.replace [k] = kw [k]
    # end def __init__

    def apply (self, root, transformer) :
        body = root.find (OOo_Tag ('office', 'body'))
        for node in body.findall ('.//' + OOo_Tag ('text', 'variable-set')) :
            name = node.get (OOo_Tag ('text', 'name'))
            if self.replace.has_key (name) :
                node.text = self.replace [name]
            elif callable (self.callback) :
                replace = self.callable (name)
                if replace :
                    node.text = replace
    # end def apply
# end class Field_Replace_Transform

class Addpagebreak_Style_Transform (Transform) :
    """This transformation adds a new ad-hoc paragraph style to the
       content part of the OOo document. This is needed to be able to
       add new page breaks to an OOo document. Adding a new page break
       is then a matter of adding an empty paragraph with the given page
       break style.

       We first look through all defined paragraph styles for
       determining a new paragraph style number. Convention is P<num>
       for paragraph styles. We search the highest number and use this
       incremented by one for the new style to insert. Then we insert
       the new style and store the resulting style name in the
       transformer under the key class_name:stylename where class_name
       is our own class name.
    """
    filename = 'content.xml'
    prio     = 80
    para     = re.compile (r'P([0-9]+)')

    def apply (self, root, transformer) :
        max_style = 0
        styles = root.find (OOo_Tag ('office', 'automatic-styles'))
        for s in styles.findall ('./' + OOo_Tag ('style', 'style')) :
            m = self.para.match (s.get (OOo_Tag ('style', 'name'), ''))
            if m :
                num = int (m.group (1))
                if num > max_style :
                    max_style = num
        stylename = 'P%d' % (max_style + 1)
        new = SubElement \
            ( styles
            , OOo_Tag ('style', 'style')
            , { OOo_Tag ('style', 'name')              : stylename
              , OOo_Tag ('style', 'family')            : 'paragraph'
              , OOo_Tag ('style', 'parent-style-name') : 'Standard'
              }
            )
        SubElement \
            ( new
            , OOo_Tag ('style', 'properties')
            , { OOo_Tag ('fo', 'break-before') : 'page' }
            )
        transformer [self.classname_prefix + 'stylename'] = stylename
    # end def apply
# end class Addpagebreak_Style_Transform

class Addpagebreak_Transform (Transform) :
    """This transformation adds a page break to the last page of the OOo
       text. This is needed, e.g., when doing mail-merge: We append a
       page break to the body and then append the next page. This
       transform needs the name of the paragraph style specifying the
       page break style. Default is to use
       'Addpagebreak_Style_Transform:stylename' as the key for
       retrieving the page style. Alternatively the page style or the
       page style key can be specified in the constructor.
    """
    filename = 'content.xml'
    prio     = 100

    def __init__ (self, stylename = None, stylekey = None) :
        self.stylename = stylename
        self.stylekey  = stylekey or 'Addpagebreak_Style_Transform:stylename'
    # end def __init__

    def apply (self, root, transformation) :
        """append to body e.g., <text:p text:style-name="P4"/>"""
        body = root.find (OOo_Tag ('office', 'body'))
        stylename = self.stylename or transformation [self.stylekey]
        SubElement \
            ( body
            , OOo_Tag ('text', 'p')
            , { OOo_Tag ('text', 'style-name') : stylename }
            )
    # end def apply
# end class Addpagebreak_Transform

if __name__ == '__main__' :
    from StringIO import StringIO
    sio = StringIO ()
    o   = OOoPy.OOoPy (infile = 'zfr.sxw', outfile = sio)
    t   = Transformer \
        ( Autoupdate_Transform ()
        , Editinfo_Transform   ()
        , Field_Replace_Transform 
            ( replace =
                { 'abo.payer.salutation' : 'Frau'
                , 'abo.payer.firstname'  : ''
                , 'abo.payer.lasttname'  : 'Musterfrau'
                }
            )
        , Addpagebreak_Style_Transform ()
        , Addpagebreak_Transform       ()
        )
    t.transform (o)
    o.close ()
    ov  = sio.getvalue ()
    f   = open ("out.sxw", "w")
    f.write (ov)
    f.close ()
