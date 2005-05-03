#!/usr/bin/env python
# -*- coding: iso-8859-1 -*-
# Copyright (C) 2005 Dr. Ralf Schlatterbeck Open Source Consulting.
# Reichergasse 131, A-3411 Weidling.
# Web: http://www.runtux.com Email: office@runtux.com
# All rights reserved
# ****************************************************************************
#
# This library is free software; you can redistribute it and/or modify
# it under the terms of the GNU Library General Public License as
# published by the Free Software Foundation; either version 2 of the
# License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Library General Public License for more details.
#
# You should have received a copy of the GNU Library General Public
# License along with this program; if not, write to the Free Software
# Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
# ****************************************************************************

import time
import re
from elementtree.ElementTree import dump, SubElement, Element
from OOoPy                   import OOoPy, autosuper
from Version                 import VERSION
from copy                    import deepcopy

files = ['content.xml', 'styles.xml', 'meta.xml', 'settings.xml']
namespaces = \
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

def OOo_Tag (namespace, name) :
    """Return combined XML tag"""
    return "{%s}%s" % (namespaces [namespace], name)
# end def OOo_Tag

class Transform (autosuper) :
    """
        Base class for individual transforms on OOo files. An individual
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
        """
            Registering with a transformer means being able to access
            variables stored in the tranformer by other transforms.
        """
        self.transformer = transformer
    # end def register

    def _varname (self, name) :
        """
            For fulfilling the naming convention of the transformer
            dictionary (every entry in this dictionary should be prefixed
            with the class name of the transform) we have this
            convenience method.
            Returns variable name prefixed with own class name.
        """
        return ":".join ((self.__class__.__name__, name))
    # end def _varname

    def set (self, variable, value) :
        """ Set variable in our transformer using naming convention. """
        self.transformer [self._varname (variable)] = value
    # end def set

    def apply (self, root) :
        raise NotImplementedError, 'derived transforms must implement "apply"'
    # end def apply

# end class Transform

class Transformer (autosuper) :
    """
        Class for applying a set of transforms to a given ooopy object.
        The transforms are applied to the specified file in priority
        order. When applying transforms we have a mechanism for
        communication of transforms. We give the transformer to the
        individual transforms as a parameter. The transforms may use the
        transformer like a dictionary for storing values and retrieving
        values left by previous transforms.
        As a naming convention each transform should use its class name
        as a prefix for storing values in the dictionary.
        >>> import Transforms
        >>> from Transforms import renumber_frames, renumber_sections \
            , renumber_tables, get_meta, set_meta
        >>> from StringIO import StringIO
        >>> sio = StringIO ()
        >>> o   = OOoPy (infile = 'test.sxw', outfile = sio)
        >>> c = o.read ('content.xml')
        >>> body = c.find (OOo_Tag ('office', 'body'))
        >>> body [-1].get (OOo_Tag ('text', 'style-name'))
        'Standard'
        >>> def cb (name) :
        ...     r = { 'street'     : 'Beispielstrasse 42'
        ...         , 'firstname'  : 'Hugo'
        ...         , 'salutation' : 'Frau'
        ...         }
        ...     if r.has_key (name) : return r [name]
        ...     return None
        ... 
        >>> p = get_meta
        >>> t = Transformer (p)
        >>> t ['a'] = 'a'
        >>> t ['a']
        'a'
        >>> p.set ('a', 'b')
        >>> t ['Attribute_Access:a']
        'b'
        >>> t   = Transformer (
        ...       Transforms.Autoupdate ()
        ...     , Transforms.Editinfo   ()  
        ...     , Transforms.Field_Replace (prio = 99, replace = cb)
        ...     , Transforms.Field_Replace
        ...         ( replace =
        ...             { 'salutation' : ''
        ...             , 'firstname'  : 'Erika'
        ...             , 'lastname'   : 'Musterfrau'
        ...             , 'country'    : 'D' 
        ...             , 'postalcode' : '00815'
        ...             , 'city'       : 'Niemandsdorf'
        ...             }
        ...         )
        ...     , Transforms.Addpagebreak_Style ()
        ...     , Transforms.Addpagebreak       ()
        ...     )
        >>> t.transform (o)
        >>> o.close ()
        >>> ov  = sio.getvalue ()
        >>> f   = open ("testout.sxw", "w")
        >>> f.write (ov)
        >>> f.close ()
        >>> o = OOoPy (infile = sio)
        >>> c = o.read ('content.xml')
        >>> body = c.find (OOo_Tag ('office', 'body'))
        >>> for node in body.findall ('.//' + OOo_Tag ('text', 'variable-set')):
        ...     name = node.get (OOo_Tag ('text', 'name'))
        ...     print name, ':', node.text
        salutation : None
        firstname : Erika
        lastname : Musterfrau
        street : Beispielstrasse 42
        country : D
        postalcode : 00815
        city : Niemandsdorf
        salutation : None
        firstname : Erika
        lastname : Musterfrau
        street : Beispielstrasse 42
        country : D
        postalcode : 00815
        city : Niemandsdorf
        >>> body [-1].get (OOo_Tag ('text', 'style-name'))
        'P2'
        >>> sio = StringIO ()
        >>> o   = OOoPy (infile = 'test.sxw', outfile = sio)
        >>> c = o.read ('content.xml')
        >>> t   = Transformer (
        ...       get_meta
        ...     , Transforms.Addpagebreak_Style ()
        ...     , Transforms.Mailmerge
        ...       ( iterator = 
        ...         ( dict (firstname = 'Erika', lastname = 'Nobody')
        ...         , dict (firstname = 'Eric',  lastname = 'Wizard')
        ...         , cb
        ...         )
        ...       )
        ...     , Transforms.Attribute_Access
        ...       ( ( renumber_frames
        ...         , renumber_sections
        ...         , renumber_tables
        ...       ) )
        ...     , set_meta
        ...     )
        >>> t.transform (o)
        >>> t [':'.join (('Set_Attribute', 'page-count'))]
        '3'
        >>> t [':'.join (('Set_Attribute', 'paragraph-count'))]
        '113'
        >>> name = t ['Addpagebreak_Style:stylename']
        >>> name
        'P2'
        >>> o.close ()
        >>> ov  = sio.getvalue ()
        >>> f   = open ("testout2.sxw", "w")
        >>> f.write (ov)
        >>> f.close ()
        >>> o = OOoPy (infile = sio)
        >>> c = o.read ('content.xml')
        >>> body = c.find (OOo_Tag ('office', 'body'))
        >>> for n in body.findall ('.//' + OOo_Tag ('text', 'p')) :
        ...     if n.get (OOo_Tag ('text', 'style-name')) == name :
        ...         print n.tag
        {http://openoffice.org/2000/text}p
        {http://openoffice.org/2000/text}p
        >>> for n in body.findall ('.//' + OOo_Tag ('text', 'variable-set')) :
        ...     if n.get (OOo_Tag ('text', 'name'), None).endswith ('name') :
        ...         name = n.get (OOo_Tag ('text', 'name'))
        ...         print name, ':', n.text
        firstname : Erika
        lastname : Nobody
        firstname : Eric
        lastname : Wizard
        firstname : Hugo
        lastname : Testman
        firstname : Erika
        lastname : Nobody
        firstname : Eric
        lastname : Wizard
        firstname : Hugo
        lastname : Testman
        >>> for n in body.findall ('.//' + OOo_Tag ('draw', 'text-box')) :
        ...     print n.get (OOo_Tag ('draw', 'name')),
        ...     print n.get (OOo_Tag ('text', 'anchor-page-number'))
        Frame1 1
        Frame2 2
        Frame3 3
        Frame4 None
        Frame5 None
        Frame6 None
        >>> for n in body.findall ('.//' + OOo_Tag ('text', 'section')) :
        ...     print n.get (OOo_Tag ('text', 'name'))
        Section1
        Section2
        Section3
        Section4
        Section5
        Section6
        Section7
        Section8
        Section9
        Section10
        Section11
        Section12
        Section13
        Section14
        Section15
        Section16
        Section17
        Section18
        >>> for n in body.findall ('.//' + OOo_Tag ('table', 'table')) :
        ...     print n.get (OOo_Tag ('table', 'name'))
        Table1
        Table2
        Table3
        >>> m = o.read ('meta.xml')
        >>> metainfo = m.find ('.//' + OOo_Tag ('meta', 'document-statistic'))
        >>> for i in 'paragraph-count', 'page-count', 'character-count' :
        ...     metainfo.get (OOo_Tag ('meta', i))
        '113'
        '3'
        '951'
        >>> o.close ()
    """
    def __init__ (self, *tf) :
        self.transforms = {}
        for t in tf :
            self.insert (t)
        self.dictionary   = {}
        self.has_key      = self.dictionary.has_key
        self.__contains__ = self.has_key
    # end def __init__

    def insert (self, transform) :
        """Insert a new transform"""
        t = transform
        if t.prio not in self.transforms :
            self.transforms [t.prio] = []
        self.transforms [t.prio].append (t)
        t.register (self)
    # end def append

    def transform (self, ooopy) :
        """
            Apply all the transforms in priority order.
            Priority order is global over all transforms.
        """
        self.trees = {}
        for f in files :
            self.trees [f] = ooopy.read (f)
            
        prios = self.transforms.keys ()
        prios.sort ()
        for p in prios :
            for t in self.transforms [p] :
                t.apply (self.trees [t.filename].getroot ())
        for e in self.trees.itervalues () :
            e.write ()
    # end def transform

    def __getitem__ (self, key) :
        return self.dictionary [key]
    # end def __getitem__

    def __setitem__ (self, key, value) :
        self.dictionary [key] = value
    # end def __setitem__
# end class Transformer
