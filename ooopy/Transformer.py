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
namespace_by_name = \
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

namespace_by_url = {}
for k,v in namespace_by_name.iteritems () :
    namespace_by_url [v] = k

def OOo_Tag (namespace, name) :
    """Return combined XML tag"""
    return "{%s}%s" % (namespace_by_name [namespace], name)
# end def OOo_Tag

def split_tag (tag) :
    """ Split tag into symbolic namespace and name part -- inverse
        operation of OOo_Tag.
    """
    ns, t = tag.split ('}')
    return (namespace_by_url [ns [1:]], t)
# end def split_tag

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
    def __init__ (self, prio = None, transformer = None) :
        if prio :
            self.prio    = prio
        self.transformer = None
        if transformer :
            self.register (transformer)
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

    def apply_all (self, trees) :
        """ Apply myself to all the files given in trees. The variable
            trees contains a dictionary of ElementTree indexed by the
            name of the OOo File.
            The standard case is that only one file (namely
            self.filename) is used.
        """
        assert (self.filename)
        self.apply (trees [self.filename].getroot ())
    # end def apply_all

    def apply (self, root) :
        """ Apply myself to the element given as root """
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
        >>> sio = StringIO ()
        >>> o   = OOoPy (infile = 'test.sxw', outfile = sio)
        >>> t   = Transformer (
        ...       get_meta
        ...     , Transforms.Concatenate ('test.sxw', 'rechng.sxw')
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
        '168'
        >>> t [':'.join (('Set_Attribute', 'character-count'))]
        '1131'
        >>> o.close ()
        >>> ov  = sio.getvalue ()
        >>> f   = open ("testout3.sxw", "w")
        >>> f.write (ov)
        >>> f.close ()
        >>> o = OOoPy (infile = sio)
        >>> c = o.read ('content.xml')
        >>> s = o.read ('styles.xml')
        >>> for n in c.findall ('./*/*') :
        ...     name = n.get (OOo_Tag ('style', 'name'))
        ...     if name :
        ...         parent = n.get (OOo_Tag ('style', 'parent-style-name'))
        ...         print '"%s", "%s"' % (name, parent)
        "Tahoma1", "None"
        "Bitstream Vera Sans", "None"
        "Tahoma", "None"
        "Nimbus Roman No9 L", "None"
        "Courier New", "None"
        "Arial Black", "None"
        "New Century Schoolbook", "None"
        "Helvetica", "None"
        "Table1", "None"
        "Table1.A", "None"
        "Table1.A1", "None"
        "Table1.E1", "None"
        "Table1.A2", "None"
        "Table1.E2", "None"
        "P1", "None"
        "fr1", "Frame"
        "fr2", "None"
        "fr3", "Frame"
        "Sect1", "None"
        "gr1", "None"
        "P2", "Standard"
        "Standard_Concat", "None"
        "Concat_P1", "Concat_Frame contents"
        "Concat_P2", "Concat_Frame contents"
        "P3", "Concat_Frame contents"
        "P4", "Concat_Frame contents"
        "P5", "Standard"
        "P6", "Standard"
        "P7", "Concat_Frame contents"
        "P8", "Concat_Frame contents"
        "P9", "Concat_Frame contents"
        "P10", "Concat_Frame contents"
        "P11", "Concat_Frame contents"
        "P12", "Concat_Frame contents"
        "P13", "Concat_Frame contents"
        "P15", "Standard"
        "P16", "Standard"
        "P17", "Standard"
        "P18", "Standard"
        "P19", "Standard"
        "P20", "Standard"
        "P21", "Standard"
        "P22", "Standard"
        "P23", "Standard"
        "T1", "None"
        "Concat_fr1", "Concat_Frame"
        "Concat_fr2", "Concat_Frame"
        "Concat_fr3", "Concat_Frame"
        "fr4", "Concat_Frame"
        "fr5", "Concat_Frame"
        "fr6", "Concat_Frame"
        "Concat_Sect1", "None"
        "N0", "None"
        "N2", "None"
        "P15_Concat", "Standard"
        >>> for n in s.findall ('./*/*') :
        ...     name = n.get (OOo_Tag ('style', 'name'))
        ...     if name :
        ...         parent = n.get (OOo_Tag ('style', 'parent-style-name'))
        ...         print '"%s", "%s"' % (name, parent)
        "Tahoma1", "None"
        "Bitstream Vera Sans", "None"
        "Tahoma", "None"
        "Nimbus Roman No9 L", "None"
        "Courier New", "None"
        "Arial Black", "None"
        "New Century Schoolbook", "None"
        "Helvetica", "None"
        "Standard", "None"
        "Text body", "Standard"
        "List", "Text body"
        "Table Contents", "Text body"
        "Table Heading", "Table Contents"
        "Caption", "Standard"
        "Frame contents", "Text body"
        "Index", "Standard"
        "Frame", "None"
        "OLE", "None"
        "Concat_Text body", "Standard"
        "Concat_List", "Concat_Text body"
        "Concat_Caption", "Standard"
        "Concat_Frame contents", "Concat_Text body"
        "Horizontal Line", "Standard"
        "Internet link", "None"
        "Visited Internet Link", "None"
        "Concat_Frame", "None"
        "Concat_OLE", "None"
        "pm1", "None"
        "Concat_pm1", "None"
        "Standard", "None"
        "Concat_Standard", "None"
        >>> for n in c.findall ('.//' + OOo_Tag ('text', 'variable-decl')) :
        ...     name = n.get (OOo_Tag ('text', 'name'))
        ...     print name
        salutation
        firstname
        lastname
        street
        country
        postalcode
        city
        date
        invoice.invoice_no
        invoice.abo.aboprice.abotype.description
        address.salutation
        address.title
        address.firstname
        address.lastname
        address.function
        address.street
        address.country
        address.postalcode
        address.city
        invoice.subscriber.salutation
        invoice.subscriber.title
        invoice.subscriber.firstname
        invoice.subscriber.lastname
        invoice.subscriber.function
        invoice.subscriber.street
        invoice.subscriber.country
        invoice.subscriber.postalcode
        invoice.subscriber.city
        invoice.period_start
        invoice.period_end
        invoice.currency.name
        invoice.amount
        invoice.subscriber.initial
        >>> for n in c.findall ('.//' + OOo_Tag ('text', 'sequence-decl')) :
        ...     name = n.get (OOo_Tag ('text', 'name'))
        ...     print name
        Illustration
        Table
        Text
        Drawing
        >>> for n in c.findall ('.//' + OOo_Tag ('text', 'p')) :
        ...     name = n.get (OOo_Tag ('text', 'style-name'))
        ...     if not name or name.startswith ('Concat') :
        ...         print ">%s<" % name
        >Concat_P1<
        >Concat_P2<
        >Concat_Frame contents<
        >>> for n in c.findall ('.//' + OOo_Tag ('draw', 'text-box')) :
        ...     attrs = 'name', 'style-name', 'z-index'
        ...     attrs = [n.get (OOo_Tag ('draw', i)) for i in attrs]
        ...     attrs.append (n.get (OOo_Tag ('text', 'anchor-page-number')))
        ...     print attrs
        ['Frame7', 'fr1', '0', '1']
        ['Frame8', 'fr1', '0', '2']
        ['Frame9', 'Concat_fr1', '0', '3']
        ['Frame10', 'Concat_fr2', '1', '3']
        ['Frame11', 'Concat_fr3', '2', '3']
        ['Frame12', 'Concat_fr1', '3', '3']
        ['Frame13', 'fr4', '4', '3']
        ['Frame14', 'fr4', '5', '3']
        ['Frame15', 'fr4', '6', '3']
        ['Frame16', 'fr4', '7', '3']
        ['Frame17', 'fr4', '8', '3']
        ['Frame18', 'fr4', '9', '3']
        ['Frame19', 'fr5', '10', '3']
        ['Frame20', 'fr4', '12', '3']
        ['Frame21', 'fr4', '13', '3']
        ['Frame22', 'fr4', '14', '3']
        ['Frame23', 'fr6', '11', '3']
        ['Frame24', 'fr4', '17', '3']
        ['Frame25', 'fr3', '2', None]
        ['Frame26', 'fr3', '2', None]
        >>> for n in c.findall ('.//' + OOo_Tag ('draw', 'rect')) :
        ...     attrs = 'style-name', 'text-style-name', 'z-index'
        ...     attrs = [n.get (OOo_Tag ('draw', i)) for i in attrs]
        ...     attrs.append (n.get (OOo_Tag ('text', 'anchor-page-number')))
        ...     print attrs
        ['gr1', 'P1', '1', '1']
        ['gr1', 'P1', '1', '2']
        >>> for n in c.findall ('.//' + OOo_Tag ('draw', 'line')) :
        ...     attrs = 'style-name', 'text-style-name', 'z-index'
        ...     attrs = [n.get (OOo_Tag ('draw', i)) for i in attrs]
        ...     print attrs
        ['gr1', 'P1', '18']
        ['gr1', 'P1', '16']
        ['gr1', 'P1', '15']
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
                t.apply_all (self.trees)
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
