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
        self.transformers = {}
        for f in files :
            self.transformers [f] = {}
        for t in ts :
            self.insert (t)
    # end def __init__

    def insert (self, t) :
        if not self.transformers [t.filename].has_key (t.prio) :
            self.transformers [t.filename][t.prio] = []
        self.transformers [t.filename][t.prio].append (t)
    # end def append

    def transform (self, ooopy) :
        for f in self.transformers.keys () :
            e = ooopy.read (f)
            root = e.getroot ()
            prios = self.transformers [f].keys ()
            prios.sort ()
            for prio in prios :
                for t in self.transformers [f][prio] :
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
        for node in root.findall (".//*") :
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
        for node in root.findall (".//" + OOo_Tag ('text', 'variable-set')) :
            name = node.get (OOo_Tag ('text', 'name'))
            if self.replace.has_key (name) :
                node.text = self.replace [name]
    # end def apply
# end class Field_Replace_Transform

class Autoupdate_Transform (Transform) :
    filename = 'settings.xml'
    prio     = 20

    def apply (self, root) :
        config = root.find ('.//' + OOo_Tag ('config', 'config-item-set'))
        for node in config.findall ('.//' + OOo_Tag ('config', 'config-item')) :
            name = node.get (OOo_Tag ('config', 'name'))
            if name == 'LinkUpdateMode' :  # update when reading
                node.text = '2'
            # update fields when reading
            if name == 'FieldAutoUpdate' or name == 'ChartAutoUpdate' :
                node.text = 'true'
    # end def apply
# end class Autoupdate_Transform

if __name__ == '__main__' :
    o = OOoPy.copy ('zfr.sxw')
    t = Transformer \
        ( Autoupdate_Transform ()
        , Editinfo_Transform   ()
        , Field_Replace_Transform 
            ( { 'abo.payer.salutation' : 'Frau'
              , 'abo.payer.firstname'  : ''
              , 'abo.payer.lasttname'  : 'Musterfrau'
              }
            )
        )
    t.transform (o)
    print o.filename
    o.close ()
