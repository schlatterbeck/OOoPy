#!/usr/bin/env python
import OOoPy
import time
from elementtree.ElementTree import dump

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
    return "{%s}%s" % (tags [tag], name)
# end def OOo_Tag

class Transform (object) :
    def __init__ (self, prio = None) :
        if prio :
            self.prio    = prio
        self.transformer = None
    # end def __init__

    def register (self, transformer) :
        self.transformer = transformer
        transformer.register (self)
    # end def register
# end class Transform

class Transformer (object) :
    def __init__ (self, *ts) :
        self.transforms = {}
        for f in files :
            self.transforms [f] = {}
        for t in ts :
            self.insert (t)
    # end def __init__

    def insert (self, t) :
        if not self.transforms [t.filename].has_key (t.prio) :
            self.transforms [t.filename][t.prio] = []
        self.transforms [t.filename][t.prio].append (t)
    # end def append

    def transform (self, ooopy) :
        for f in self.transforms.keys () :
            e = ooopy.read (f)
            root = e.getroot ()
            prios = self.transforms [f].keys ()
            prios.sort ()
            for prio in prios :
                for t in self.transforms [f][prio] :
                    t.apply (root)
            e.write ()
    # end def transform
# end class Transformer

class Editinfo_Transform (Transform) :
    filename = 'meta.xml'
    prio     = 20
    replace = \
    { OOo_Tag ('meta', 'generator')        : 'OOoPy field replacement'
    , OOo_Tag ('dc',   'date')             : time.strftime ('%Y-%m-%dT%H:%M:%S')
    , OOo_Tag ('meta', 'editing-cycles')   : '0'
    , OOo_Tag ('meta', 'editing-duration') : 'PT0M0S'
    }

    def apply (self, root) :
        for node in root.findall (OOo_Tag ('office', 'meta') + '/*') :
            if self.replace.has_key (node.tag) :
                node.text = self.replace [node.tag]
    # end def apply
# end class Editinfo_Transform

class Field_Replace_Transform (Transform) :
    filename = 'content.xml'
    prio     = 100

    def __init__ (self, prio = None, replace = None, **kw) :
        Transform (prio)
        self.replace = replace or {}
        for k in kw.keys () :
            self.replace [k] = kw [k]
    # end def __init__

    def apply (self, root) :
        body = root.find (OOo_Tag ('office', 'body'))
        for node in body.findall ('.//' + OOo_Tag ('text', 'variable-set')) :
            name = node.get (OOo_Tag ('text', 'name'))
            if self.replace.has_key (name) :
                node.text = self.replace [name]
    # end def apply
# end class Field_Replace_Transform

class Autoupdate_Transform (Transform) :
    filename = 'settings.xml'
    prio     = 20

    def apply (self, root) :
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

class Addpagebreak_Transform (Transform) :
    # Add a pagebreak paragraph style
    # <style:style style:name="P4" style:family="paragraph"
    # style:parent-style-name="Standard"><style:properties
    # fo:font-size="10pt"
    # style:font-size-asian="10pt" style:font-size-complex="10pt"
    # fo:break-before="page"/></style:style>
    # then add the pagebreak as the last element in body
    # <text:p text:style-name="P4"/>
# end class Addpagebreak_Transform

if __name__ == '__main__' :
    o = OOoPy.copy ('zfr.sxw')
    t = Transformer \
        ( Autoupdate_Transform ()
        , Editinfo_Transform   ()
        , Field_Replace_Transform 
            ( replace =
                { 'abo.payer.salutation' : 'Frau'
                , 'abo.payer.firstname'  : ''
                , 'abo.payer.lasttname'  : 'Musterfrau'
                }
            )
        )
    t.transform (o)
    print o.filename
    o.close ()
