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
from elementtree.ElementTree import dump, SubElement, Element, tostring
from OOoPy                   import OOoPy, autosuper
from Transformer             import files, split_tag, OOo_Tag, Transform
from Version                 import VERSION
from copy                    import deepcopy

# counts in meta.xml
meta_counts = \
    ( 'character-count', 'image-count', 'object-count', 'page-count'
    , 'paragraph-count', 'table-count', 'word-count'
    )

def _body () :
    """
        We use the body element as a container for various
        transforms...
    """
    return Element (OOo_Tag ('office', 'body'))
# end def _body

class Access_Attribute (autosuper) :
    """ For performance reasons we do not specify a separate transform
        for each attribute-read or -change operation. Instead we define
        all the attribute accesses we want to perform as objects that
        follow the attribute access api and apply them all using an
        Attribute_Access in one go.
    """

    def __init__ (self, key = None, prefix = None, ** kw) :
        self.__super.__init__ (key = key, prefix = prefix, **kw)
        self.key = key
        if key :
            if not prefix :
                prefix   = self.__class__.__name__
            self.key = ':'.join ((prefix, key))
    # end def __init__

    def register (self, transformer) :
        self.transformer = transformer
    # end def register

    def use_value (self, oldval = None) :
        """ Can change the given value by returning the new value. If
            returning None or oldval the attribute stays unchanged.
        """
        raise NotImplementedError, "use_value must be defined in derived class"
    # end def use_value

# end class Access_Attribute

class Get_Attribute (Access_Attribute) :
    """ An example of not changing an attribute but only storing the
        value in the transformer
    """

    def __init__ (self, tag, attr, key, transform = None, ** kw) :
        self.__super.__init__ (key = key, **kw)
        self.tag        = tag
        self.attribute  = attr
        self.transform  = transform
    # end def __init__

    def use_value (self, oldval = None) :
        self.transformer [self.key] = oldval
        return None
    # end def use_value

# end def Get_Attribute

class Get_Max (Access_Attribute) :
    """ Get the maximum value of an attribute """

    def __init__ (self, tag, attr, key, transform = None, ** kw) :
        self.__super.__init__ (key = key, **kw)
        self.tag        = tag
        self.attribute  = attr
        self.transform  = transform
    # end def __init__

    def register (self, transformer) :
        self.__super.register (transformer)
        self.transformer [self.key] = -1
    # end def register

    def use_value (self, oldval = None) :
        if  self.transformer [self.key] < oldval :
            self.transformer [self.key] = oldval
        return None
    # end def use_value

# end def Get_Max

class Renumber (Access_Attribute) :
    """ Specifies a renumbering transform. OOo has a 'name' attribute
        for several different tags, e.g., tables, frames, sections etc.
        These names must be unique in the whole document. OOo itself
        solves this by appending a unique number to a basename for each
        element, e.g., sections are named 'Section1', 'Section2', ...
        Renumber transforms can be applied to correct the numbering
        after operations that destroy the unique numbering, e.g., after
        a mailmerge where the same document is repeatedly appended.
    """

    def __init__ (self, tag, name = None, attr = None, start = 1) :
        self.__super.__init__ ()
        tag_ns, tag_name = split_tag (tag)
        self.tag         = tag
        self.name        = name or tag_name [0].upper () + tag_name [1:]
        self.num         = start
        self.attribute   = attr or OOo_Tag (tag_ns, 'name')
    # end def __init__

    def use_value (self, oldval = None) :
        name = "%s%d" % (self.name, self.num)
        self.num += 1
        return name
    # end def use_value

# end class Renumber

class Set_Attribute (Access_Attribute) :
    """
        Similar to the renumbering transform in that we are assigning
        new values to some attributes. But in this case we give keys
        into the Transformer dict to replace some tag attributes.
    """

    def __init__ \
        ( self
        , tag
        , attr
        , key       = None
        , transform = None
        , value     = None
        , oldvalue  = None
        , ** kw
        ) :
        self.__super.__init__ (key = key, ** kw)
        self.tag        = tag
        self.attribute  = attr
        self.transform  = transform
        self.value      = value
        self.oldvalue   = oldvalue
    # end def __init__

    def use_value (self, oldval) :
        if oldval is None :
            return None
        if self.oldvalue and oldval != self.oldvalue :
            return None
        if self.key and self.transformer.has_key (self.key) :
            return str (self.transformer [self.key])
        return self.value
    # end def use_value

# end class Set_Attribute

def set_attributes_from_dict (tag, attr, d) :
    """ Convenience function: iterate over a dict and return a list of
        Set_Attribute objects specifying replacement of attributes in
        the dictionary
    """
    return [Set_Attribute (tag, attr, oldvalue = k, value = v)
            for k,v in d.iteritems ()
           ]
# end def set_attributes_from_dict

class Reanchor (Access_Attribute) :
    """
        Similar to the renumbering transform in that we are assigning
        new values to some attributes. But in this case we want to
        relocate objects that are anchored to a page.
    """

    def __init__ (self, offset, tag, attr = None) :
        self.__super.__init__ ()
        self.offset     = int (offset)
        self.tag        = tag
        self.attribute  = attr or OOo_Tag ('text', 'anchor-page-number')
    # end def __init__

    def use_value (self, oldval) :
        if oldval is None :
            return oldval
        return "%d" % (int (oldval) + self.offset)
    # end def use_value

# end class Reanchor

#
# general transforms applicable to several .xml files
#

class Attribute_Access (Transform) :
    """
        Read or Change attributes in an OOo document.
        Can be used for renumbering, moving anchored objects, etc.
        Expects a list of attribute changer objects that follow the
        attribute changer API. This API is very simple:

        - Member function "use_value" returns the new value of an
          attribute, or if unchanged the old value
        - The attribute "tag" gives the tag for an element we are
          searching
        - The attribute "attribute" gives the name of the attribute we
          want to read or change.
        For examples of the attribute changer API, see Renumber and
        Reanchor above.
    """
    filename = 'content.xml'
    prio     = 110

    def __init__ (self, attrchangers, filename = None, ** kw) :
        self.filename     = filename or self.filename
        self.attrchangers = {}
        # allow several changers for a single tag
        self.attrchangers [None] = []
        for r in attrchangers :
            if r.tag not in self.attrchangers :
                self.attrchangers [r.tag] = []
            self.attrchangers [r.tag].append (r)
        self.__super.__init__ (** kw)
    # end def __init__

    def register (self, transformer) :
        """ Register transformer with all attrchangers. """
        self.__super.register (transformer)
        for a in self.attrchangers.itervalues () :
            for r in a :
                r.register (transformer)
    # end def register

    def apply (self, root) :
        """ Search for all tags for which we renumber and replace name """
        for n in [root] + root.findall ('.//*') :
            changers = \
                self.attrchangers [None] + self.attrchangers.get (n.tag, [])
            for r in changers :
                nval = r.use_value (n.get (r.attribute))
                if nval is not None :
                    n.set (r.attribute, nval)
    # end def apply

# end class Attribute_Access

#
# meta.xml transforms
#

class Editinfo (Transform) :
    """
        This is an example of modifying OOo meta info (edit information,
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

    def apply (self, root) :
        for node in root.findall (OOo_Tag ('office', 'meta') + '/*') :
            if self.replace.has_key (node.tag) :
                node.text = self.replace [node.tag]
    # end def apply
# end class Editinfo

#
# settings.xml transforms
#

class Autoupdate (Transform) :
    """
        This is an example of modifying OOo settings. We set some of the
        AutoUpdate configuration items in OOo to true. We also specify
        that links should be updated when reading.

        This was originally intended to make OOo correctly display fields
        if they were changed with the Field_Replace below
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
# end class Autoupdate

#
# content.xml transforms
#

class Field_Replace (Transform) :
    """
        Takes a dict of replacement key-value pairs. The key is the name
        of a variable in OOo. Additional replacement key-value pairs may
        be specified in ** kw. Alternatively a callback mechanism for
        variable name lookups is provided. The callback function is
        given the name of a variable in OOo and is expected to return
        the replacement value or None if the variable value should not
        be replaced.
    """
    filename = 'content.xml'
    prio     = 100

    def __init__ (self, prio = None, replace = None, ** kw) :
        """ replace is something behaving like a dict or something
            callable for name lookups
        """
        self.__super.__init__ (prio)
        self.replace  = replace or {}
        self.dict     = kw
    # end def __init__

    def apply (self, root) :
        body = root
        if body.tag != OOo_Tag ('office', 'body') :
            body = body.find (OOo_Tag ('office', 'body'))
        for tag in 'variable-set', 'variable-get', 'variable-input' :
            for node in body.findall ('.//' + OOo_Tag ('text', tag)) :
                name = node.get (OOo_Tag ('text', 'name'))
                if callable (self.replace) :
                    replace = self.replace (name)
                    if replace :
                        node.text = replace
                elif name in self.replace :
                    node.text = self.replace [name]
                elif name in self.dict :
                    node.text = self.dict    [name]
    # end def apply
# end class Field_Replace

class Addpagebreak_Style (Transform) :
    """
        This transformation adds a new ad-hoc paragraph style to the
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
    prio     = 30
    para     = re.compile (r'P([0-9]+)')

    def apply (self, root) :
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
            , { OOo_Tag ('fo', 'break-after') : 'page' }
            )
        self.set ('stylename', stylename)
    # end def apply
# end class Addpagebreak_Style

class Addpagebreak (Transform) :
    """
        This transformation adds a page break to the last page of the OOo
        text. This is needed, e.g., when doing mail-merge: We append a
        page break to the body and then append the next page. This
        transform needs the name of the paragraph style specifying the
        page break style. Default is to use
        'Addpagebreak_Style:stylename' as the key for
        retrieving the page style. Alternatively the page style or the
        page style key can be specified in the constructor.
    """
    filename = 'content.xml'
    prio     = 50

    def __init__ (self, stylename = None, stylekey = None, ** kw) :
        self.__super.__init__ (** kw)
        self.stylename = stylename
        self.stylekey  = stylekey or 'Addpagebreak_Style:stylename'
    # end def __init__

    def apply (self, root) :
        """append to body e.g., <text:p text:style-name="P4"/>"""
        body = root
        if body.tag != OOo_Tag ('office', 'body') :
            body = body.find (OOo_Tag ('office', 'body'))
        stylename = self.stylename or self.transformer [self.stylekey]
        SubElement \
            ( body
            , OOo_Tag ('text', 'p')
            , { OOo_Tag ('text', 'style-name') : stylename }
            )
    # end def apply
# end class Addpagebreak

class _Body_Concat (Transform) :
    """ Various methods for modifying the body split into various pieces
        that have to keep sequence in order to not confuse OOo.
    """
    ooo_sections = \
        [ { OOo_Tag ('text', 'variable-decls') : 1
          , OOo_Tag ('text', 'sequence-decls') : 1
          }
        , { OOo_Tag ('draw', 'text-box')       : 1
          , OOo_Tag ('draw', 'rect')           : 1
          }
        ]

    def _divide (self, body) :
        """ Divide self.copy into parts that must keep their sequence.
            We use another body tag for storing the parts...
            Side-effect that self.copyparts is set is intended.
        """
        self.copyparts = _body ()
        self.copyparts.append (_body ())
        l = len (self.ooo_sections)
        idx = 0
        for e in body :
            if idx < l :
                if e.tag not in self.ooo_sections [idx] :
                    self.copyparts.append (_body ())
                    idx += 1
            self.copyparts [-1].append (e)
        declarations = self.copyparts [0]
        del self.copyparts [0]
        return declarations
    # end def _divide

    def divide_body (self, root) :
        cont       = root
        if cont.tag != OOo_Tag ('office', 'document-content') :
            cont   = root.find  (OOo_Tag ('office', 'document-content'))
        body       = cont.find  (OOo_Tag ('office', 'body'))
        idx        = cont [:].index (body)
        self.body  = cont [idx] = _body ()
        self.declarations = self._divide (body)
        self.bodyparts    = self.copyparts
    # end def divide_body

    def append_declarations (self) :
        for e in self.declarations :
            self.body.append (e)
    # end def append_declarations

    def append_to_body (self, cp) :
        for i in range (len (self.bodyparts)) :
            for j in cp [i] :
                self.bodyparts [i].append (j)
    # end def append_to_body

    def assemble_body (self) :
        for p in self.bodyparts :
            for e in p :
                self.body.append (e)
    # end def assemble_body

    def _get_meta (self, var, classname = 'Get_Attribute', prefix = "") :
        """ get page- and paragraph-count etc. meta-info """
        return int (self.transformer [':'.join ((classname, prefix + var))])
    # end def _get_meta

    def _set_meta (self, var, value, classname = 'Set_Attribute', prefix = "") :
        """ set page- and paragraph-count etc. meta-info """
        self.transformer [':'.join ((classname, prefix + var))] = str (value)
    # end def _set_meta
# end class _Body_Concat

class Mailmerge (_Body_Concat) :
    """
        This transformation is used to create a mailmerge document using
        the current document as the template. In the constructor we get
        an iterator that provides a data set for each item in the
        iteration. Elements the iterator has to provide are either
        something that follows the Mapping Type interface (it looks like
        a dict) or something that is callable and can be used for
        name-value lookups.

        A precondition for this transform is the application of the
        Addpagebreak_Style to guarantee that we know the style
        for adding a page break to the current document. Alternatively
        the stylename (or the stylekey if a different name should be used
        for lookup in the current transformer) can be given in the
        constructor.
    """
    filename = 'content.xml'
    prio     = 60

    def __init__ \
        (self, iterator, stylename = None, stylekey = None, ** kw) :
        self.__super.__init__ (** kw)
        self.iterator  = iterator
        self.stylename = stylename
        self.stylekey  = stylekey
    # end def __init__

    def apply (self, root) :
        """
            Copy old body, create new empty one and repeatedly append the
            new body.
        """
        pb = Addpagebreak \
            ( stylename   = self.stylename
            , stylekey    = self.stylekey
            , transformer = self.transformer
            )
        zi = Attribute_Access \
            ( (Get_Max (None, OOo_Tag ('draw', 'z-index'), 'z-index'),)
            , transformer = self.transformer
            )
        zi.apply (root)

        pagecount  = self._get_meta ('page-count')
        z_index    = self._get_meta ('z-index', classname = 'Get_Max') + 1
        ra         = Attribute_Access \
            (( Reanchor (pagecount, OOo_Tag ('draw', 'text-box'))
            ,  Reanchor (pagecount, OOo_Tag ('draw', 'rect'))
            ,  Reanchor (z_index, None, OOo_Tag ('draw', 'z-index'))
            ))
        self.divide_body (root)
        self.bodyparts = [_body () for i in self.copyparts]

        count = 0
        for i in self.iterator :
            count += 1
            fr = Field_Replace (replace = i, transformer = self.transformer)
            # add page break only to non-empty body
            # reanchor only after the first mailmerge
            if self.body :
                pb.apply (self.bodyparts [-1])
                ra.apply (self.copyparts)
            else :
                self.append_declarations ()
            cp = deepcopy (self.copyparts)
            fr.apply (cp)
            self.append_to_body (cp)
        # new page-count:
        for i in meta_counts :
            self._set_meta (i, count * self._get_meta (i))
        # we have added count-1 paragraphs, because each page-break is a
        # paragraph.
        p = 'paragraph-count'
        self._set_meta \
            (p, self._get_meta (p, classname = 'Set_Attribute') + (count - 1))
        self.assemble_body ()
    # end def apply
# end class Mailmerge

_stylename  = OOo_Tag ('style', 'name')
_parentname = OOo_Tag ('style', 'parent-style-name')
_textname   = OOo_Tag ('text',  'name')

def tree_serialise (element, prefix = '') :
    """ Serialise a style-element of an OOo document (e.g., a
        style:font-decl, style:default-style, etc declaration).
        We remove the name of the style and return something that is a
        representation of the style element which can be used as a
        dictionary key.
        The serialisation format is a tuple containing the tag as the
        first item, the attributes (as key,value pairs returned by
        items()) as the second item and the following items are
        serialisations of children.
    """
    attr = dict (element.attrib)
    if _stylename  in attr : del attr [_stylename]
    attr = attr.items ()
    attr.sort ()
    attr = tuple (attr)
    serial = [prefix + element.tag, attr]
    for e in element :
        serial.append (tree_serialise (e, prefix))
    return tuple (serial)
# end def tree_serialise

class Concatenate (_Body_Concat) :
    """
        This transformation is used to create a new document from a
        concatenation of several documents.  In the constructor we get a
        list of documents to append to the master document.
    """
    prio     = 80
    style_containers = \
      { OOo_Tag ('office', 'font-decls')       : 1
      , OOo_Tag ('office', 'styles')           : 1
      , OOo_Tag ('office', 'automatic-styles') : 1
      , OOo_Tag ('office', 'master-styles')    : 1
      }
    # Cross-references in OOo document:
    # 'attribute' references another element with 'tag'.
    # If attribute names change, we must replace references, too.
    #   attribute                             : tag
    ref_attrs = \
      { _parentname                           : OOo_Tag ('style', 'style')
      , OOo_Tag ('style', 'master-page-name') : OOo_Tag ('style', 'master-page')
      , OOo_Tag ('style', 'page-master-name') : OOo_Tag ('style', 'page-master')
      , OOo_Tag ('text',  'style-name')       : OOo_Tag ('style', 'style')
      , OOo_Tag ('draw',  'style-name')       : OOo_Tag ('style', 'style')
      , OOo_Tag ('draw',  'text-style-name')  : OOo_Tag ('style', 'style')
      }
    stylefiles = ['styles.xml', 'content.xml']
    oofiles    = stylefiles + ['meta.xml']

    body_decl_sections = ['variable-decl', 'sequence-decl']

    def __init__ \
        (self, * docs, ** kw) :
        self.__super.__init__ (** kw)
        self.docs       = [OOoPy (infile = doc) for doc in docs]
    # end def __init__

    def apply_all (self, trees) :
        self.serialised = {}
        self.stylenames = {}
        self.namemaps   = [{}]
        for s in self.ref_attrs.itervalues () :
            self.namemaps [0][s] = {}
        self.body_decls = {}
        for s in self.body_decl_sections :
            self.body_decls [s] = {}
        self.trees      = {}
        for f in self.oofiles :
            self.trees [f] = [trees [f].getroot ()]
        self.sections   = {}
        for f in self.stylefiles :
            self.sections [f] = {}
            for node in self.trees [f][0] :
                self.sections [f][node.tag] = node
        for d in self.docs :
            self.namemaps.append ({})
            for s in self.ref_attrs.itervalues () :
                self.namemaps [-1][s] = {}
            for f in self.oofiles :
                self.trees [f].append (d.read (f).getroot ())
        # append a pagebreak style, will be optimized away if duplicate
        pbs = Addpagebreak_Style (transformer = self.transformer)
        pbs.apply (self.trees ['content.xml'][0])
        get_attr = []
        for attr in meta_counts :
            a = OOo_Tag ('meta', attr)
            t = OOo_Tag ('meta', 'document-statistic')
            get_attr.append (Get_Attribute (t, a, 'concat-' + attr))
        zi = Attribute_Access \
            ( (Get_Max (None, OOo_Tag ('draw', 'z-index'), 'z-index'),)
            , transformer = self.transformer
            )
        zi.apply (self.trees ['content.xml'][0])
        self.zi = Attribute_Access \
            ( (Get_Max (None, OOo_Tag ('draw', 'z-index'), 'concat-z-index'),)
            , transformer = self.transformer
            )
        self.getmeta = Attribute_Access \
            (get_attr, filename = 'meta.xml', transformer = self.transformer)
        self.pbname = self.transformer \
            [':'.join (('Addpagebreak_Style', 'stylename'))]
        for s in self.trees ['styles.xml'][0].findall \
            ('.//' + OOo_Tag ('style', 'default-style')) :
            if s.get (OOo_Tag ('style', 'family')) == 'paragraph' :
                default_style = s
                break
        self.default_properties = default_style.find \
            ('./' + OOo_Tag ('style', 'properties'))
        self.set_pagestyle ()
        for f in 'styles.xml', 'content.xml' :
            self.style_merge (f)
        self.body_concat ()
    # end def apply_all

    def _attr_rename (self, idx) :
        r = sum \
            ( [ set_attributes_from_dict (None, k, self.namemaps [idx][v])
                for k,v in self.ref_attrs.iteritems ()
              ]
            , []
            )
        return Attribute_Access (r, transformer = self.transformer)
    # end def _attr_rename

    def body_concat (self) :
        count = {}
        for i in meta_counts :
            count [i] = self._get_meta (i)
        count ['z-index'] = self._get_meta \
            ('z-index', classname = 'Get_Max') + 1
        pb   = Addpagebreak \
            (stylename = self.pbname, transformer = self.transformer)
        self.divide_body (self.trees ['content.xml'][0])
        self.body_decl (self.declarations, append = 0)
        for idx in range (1, len (self.docs) + 1) :
            meta    = self.trees ['meta.xml'][idx]
            content = self.trees ['content.xml'][idx]
            body    = content.find (OOo_Tag ('office', 'body'))
            self.getmeta.apply (meta)
            self.zi.apply      (body)

            ra = Attribute_Access \
              (( Reanchor (count ['page-count'], OOo_Tag ('draw', 'text-box'))
              ,  Reanchor (count ['page-count'], OOo_Tag ('draw', 'rect'))
              ,  Reanchor (count ['z-index'], None, OOo_Tag ('draw', 'z-index'))
              ))
            for i in meta_counts :
                count [i] += self._get_meta (i, prefix = 'concat-')
            count ['paragraph-count'] += 1
            count ['z-index'] += self._get_meta \
                ('z-index', classname = 'Get_Max', prefix = 'concat-') + 1
            namemap = self.namemaps [idx][OOo_Tag ('style', 'style')]
            tr      = self._attr_rename (idx)
            pb.apply (self.bodyparts [-1])
            tr.apply (content)
            ra.apply (content)
            declarations = self._divide (body)
            self.body_decl (declarations)
            self.append_to_body (self.copyparts)
        self.append_declarations ()
        self.assemble_body       ()
        for i in meta_counts :
            self._set_meta (i, count [i])
    # end def body_concat

    def body_decl (self, decl_section, append = 1) :
        for sect in self.body_decl_sections :
            s = self.declarations.find ('.//' + OOo_Tag ('text', sect + 's'))
            d = self.body_decls [sect]
            t = OOo_Tag ('text', sect)
            for n in decl_section.findall ('.//' + t) :
                name = n.get (_textname)
                if name not in d :
                    if append and s : s.append (n)
                    d [name] = 1
    # end def body_decl

    def merge_defaultstyle (self, default_style, node) :
        assert default_style is not None
        assert node is not None
        proppath = './' + OOo_Tag ('style', 'properties')
        defprops = default_style.find (proppath)
        props    = node.find          (proppath)
        if props is None :
            props = Element (OOo_Tag ('style', 'properties'))
        for k,v in defprops.attrib.iteritems () :
            if self.default_properties.get (k) != v and not props.get (k) :
                if k == OOo_Tag ('style', 'tab-stop-distance') :
                    stps = SubElement (props, OOo_Tag ('style', 'tab-stops'))
                    l    = float (v [:-2])
                    unit = v [-2:]
                    for ts in range (10) :
                        SubElement \
                            ( stps
                            , OOo_Tag ('style', 'tab-stop')
                            , { OOo_Tag ('style', 'position') 
                                : '%s%s' % (l * (ts + 1), unit)
                              }
                            )
                else :
                    props.set (k,v)
        if len (props) or props.attrib :
            node.append (props)
    # end def merge_defaultstyle

    def _newname (self, key, oldname) :
        stylenum = 0
        if (key, oldname) not in self.stylenames :
            self.stylenames [(key, oldname)] = 1
            return oldname
        newname = basename = 'Concat_%s' % oldname
        while (key, newname) in self.stylenames :
            stylenum += 1
            newname = '%s%d' % (basename, stylenum)
        self.stylenames [(key, newname)] = 1
        return newname
    # end def _newname

    def set_pagestyle (self) :
        """ For all documents: search for the first paragraph of the body
            and get its style. Modify this style to include a reference
            to the default page-style if it doesn't contain a reference
            to a page style. Insert the new style into the list of
            styles and modify the first paragraph to use the new page
            style.
            This procedure is necessary to make appended documents use
            their page style instead of the master page style of the
            first document.
            FIXME: We should search the style hierarchy backwards for
            the style of the first paragraph to check if there is a
            reference to a page-style somewhere and not override the
            page-style in this case. Otherwise appending complex
            documents that use a different page-style for the first page
            will not work if the page style is referenced in a style
            from which the first paragraph style derives.
        """
        for idx in range (1, len (self.docs) + 1) :
            croot  = self.trees  ['content.xml'][idx]
            sroot  = self.trees  ['styles.xml'] [idx]
            body   = croot.find (OOo_Tag ('office', 'body'))
            para   = body.find  ('./' + OOo_Tag ('text', 'p'))
            tsn    = OOo_Tag ('text', 'style-name')
            sname  = para.get   (tsn)
            styles = croot.find (OOo_Tag ('office', 'automatic-styles'))
            ost    = sroot.find (OOo_Tag ('office', 'styles'))
            mst    = sroot.find (OOo_Tag ('office', 'master-styles'))
            assert mst
            assert mst [0].tag == OOo_Tag ('style', 'master-page')
            master = mst [0].get (_stylename)
            mpn    = OOo_Tag ('style', 'master-page-name')
            stytag = OOo_Tag ('style', 'style')
            style  = None
            for s in styles :
                if s.tag == stytag :
                    # Explicit references to default style converted to
                    # explicit references to new page style.
                    if s.get (mpn) == '' :
                        s.set (mpn, master)
                    if s.get (_stylename) == sname :
                        style = s
            if not style :
                for s in ost :
                    if s.tag == stytag and s.get (_stylename) == sname :
                        style = s
                        break
            assert style is not None
            if not style.get (mpn) :
                newstyle = deepcopy (style)
                # Don't register with newname: will be rewritten later
                # when appending. We assume that an original doc does
                # not already contain a style with _Concat suffix.
                newname = sname + '_Concat'
                para.set (tsn, newname)
                newstyle.set (OOo_Tag ('style', 'name'), newname)
                newstyle.set (mpn,                       master)
                styles.append (newstyle)
    # end def set_pagestyle

    def style_merge (self, oofile) :
        """ Loop over all the docs in our document list and look up the
            styles there. If a style matches an existing style in the
            original document, register the style name for later
            transformation if the style name in the original document
            does not match the style name in the appended document.  If
            no match is found, append style to master document and add
            to serialisation. If the style name already exists in the
            master document, a new style name is created. Names of
            parent styles are changed when appending -- this means that
            parent style names already have to be defined earlier in the
            document.

            If there is a reference to a parent style that is not yet
            defined, and the parent style is defined later, it is
            already too late, so an assertion is raised in this case.
            OOo seems to ensure declaration order of dependent styles,
            so this should not be a problem.
        """
        _fontdcl = OOo_Tag ('office', 'font-decls')
        for idx in range (len (self.trees [oofile])) :
            namemap = self.namemaps [idx]
            root    = self.trees    [oofile][idx]
            delnode = []
            for node in root :
                if node.tag not in self.style_containers :
                    continue
                prefix = ''
                if node.tag == OOo_Tag ('office', 'font-decls') :
                    prefix = oofile
                nodeidx = -1
                default_style = None
                for n in node :
                    if  (   n.tag == OOo_Tag ('style', 'default-style')
                        and n.get (OOo_Tag ('style', 'family')) == 'paragraph'
                        ) :
                        default_style = n
                    name     = n.get (_stylename, None)
                    nodeidx += 1
                    if not name : continue
                    if  (   idx != 0
                        and name == 'Standard'
                        and n.get (OOo_Tag ('style', 'class'))  == 'text'
                        and n.get (OOo_Tag ('style', 'family')) == 'paragraph'
                        ) :
                        self.merge_defaultstyle (default_style, n)
                    key = prefix + n.tag
                    if key not in namemap : namemap [key] = {}
                    tr = self._attr_rename (idx)
                    tr.apply (n)
                    sn  = tree_serialise (n, prefix)
                    if sn in self.serialised :
                        newname = self.serialised [sn]
                        if name != newname :
                            assert \
                                (  name not in namemap [key]
                                or namemap [key][name] == newname
                                )
                            namemap [key][name] = newname
                            # optimize original doc: remove duplicate styles
                            if  not idx and node.tag != _fontdcl :
                                delnode.append (nodeidx)
                    else :
                        newname = self._newname (key, name)
                        self.serialised [sn] = newname
                        if newname != name :
                            n.set (_stylename, newname)
                            namemap [key][name] = newname
                        if idx != 0 :
                            self.sections [oofile][node.tag].append (n)
            assert not delnode or not idx
            delnode.reverse ()
            for i in delnode :
                del node [i]
    # end style_register
            
# end class Concatenate

renumber_frames   = Renumber (OOo_Tag ('draw',  'text-box'), 'Frame')
renumber_sections = Renumber (OOo_Tag ('text',  'section'))
renumber_tables   = Renumber (OOo_Tag ('table', 'table'))

renumber_all      = Attribute_Access \
    ( ( renumber_frames
      , renumber_sections
      , renumber_tables
    ) )

# used to have a separate Pagecount transform -- generalized to get
# some of the meta information using an Attribute_Access transform
# and set the same information later after possibly being updated by
# other transforms. We use another naming convention here for storing
# the info retrieved from the OOo document: We use the attribute name in
# the meta-information to store (and later retrieve) the information.

get_attr = []
set_attr = []
for attr in meta_counts :
    a = OOo_Tag ('meta', attr)
    t = OOo_Tag ('meta', 'document-statistic')
    get_attr.append (Get_Attribute (t, a, attr))
    set_attr.append (Set_Attribute (t, a, attr))
get_meta = Attribute_Access (get_attr, prio =  20, filename = 'meta.xml')
set_meta = Attribute_Access (set_attr, prio = 120, filename = 'meta.xml')
