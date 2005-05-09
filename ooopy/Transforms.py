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
from Transformer             import files, split_tag, OOo_Tag, Transform
from Version                 import VERSION
from copy                    import deepcopy

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
        for n in root.findall ('.//*') :
            if n.tag in self.attrchangers :
                for r in self.attrchangers [n.tag] + self.attrchangers [None] :
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
    sections = \
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
        l = len (self.sections)
        idx = 0
        for e in body :
            if idx < l :
                if e.tag not in self.sections [idx] :
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

    def _get_meta (self, var) :
        """ get page- and paragraph-count etc. meta-info """
        return int (self.transformer [':'.join (('Get_Attribute', var))])
    # end def _get_meta

    def _set_meta (self, var, value) :
        """ set page- and paragraph-count etc. meta-info """
        self.transformer [':'.join (('Set_Attribute', var))] = str (value)
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
        pagecount  = self._get_meta ('page-count')
        ra         = Attribute_Access \
            (( Reanchor (pagecount, OOo_Tag ('draw', 'text-box'))
            ,  Reanchor (pagecount, OOo_Tag ('draw', 'rect'))
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
        for i in 'page-count', 'character-count' :
            self._set_meta (i, count * self._get_meta (i))
        # we have added count-1 paragraphs, because each page-break is a
        # paragraph.
        pars = self._get_meta ('paragraph-count') * count + (count - 1)
        self._set_meta ('paragraph-count', pars)
        self.assemble_body ()
    # end def apply
# end class Mailmerge

_stylename  = OOo_Tag ('style', 'name')
_parentname = OOo_Tag ('style', 'parent-style-name')
_textname   = OOo_Tag ('text',  'name')

def tree_serialise (element) :
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
    attr = element.attrib
    if _stylename in attr or _parentname in attr :
        attr = dict (attr)
        if _stylename  in attr : del attr [_stylename]
        if _parentname in attr : del attr [_parentname]
    attr = attr.items ()
    attr.sort ()
    attr = tuple (attr)
    serial = [element.tag, attr]
    for e in element :
        serial.append (tree_serialise (e))
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

    def __init__ \
        (self, * docs, ** kw) :
        self.__super.__init__ (** kw)
        self.docs       = [OOoPy (infile = doc) for doc in docs]
        self.serialised = {}
        self.stylenames = {}
    # end def __init__

    def apply_all (self, trees) :
        self.trees      = trees
        self.stylemaps  = []
        self.treemaps   = {'content.xml' : [], 'styles.xml' : []}
        self.containers = {'content.xml' : {}, 'styles.xml' : {}}
        content = self.trees ['content.xml'].getroot ()
        for d in self.docs :
            self.stylemaps.append ({})
            for f in 'styles.xml', 'content.xml' :
                self.treemaps [f].append (d.read (f).getroot ())
        # append a pagebreak style, will be optimized away if duplicate
        pbs = Addpagebreak_Style (transformer = self.transformer)
        pbs.apply (content)
        get_attr = []
        for attr in 'character-count', 'page-count', 'paragraph-count' :
            a = OOo_Tag ('meta', attr)
            t = OOo_Tag ('meta', 'document-statistic')
            get_attr.append (Get_Attribute (t, a, 'concat' + attr))
        self.getmeta = Attribute_Access \
            (get_attr, filename = 'meta.xml', transformer = self.transformer)
        self.pbname = self.transformer \
            [':'.join (('Addpagebreak_Style', 'stylename'))]
        for f in 'styles.xml', 'content.xml' :
            self.style_register (f)
        for f in 'styles.xml', 'content.xml' :
            self.style_concat   (f)
        self.body_concat (content)
    # end def apply_all

    def register_decls (self) :
        self.decl = {}
        for decl in 'variable-decl', 'sequence-decl' :
            d = self.decl [decl] = {}
            t = OOo_Tag ('text', decl)
            for n in self.declarations.findall ('.//' + t) :
                d [n.get (_textname)] = 1
    # end def register_decls

    def update_decls (self, decls) :
        for decl in 'variable-decl', 'sequence-decl' :
            sect = self.declarations.find ('.//' + OOo_Tag ('text', decl + 's'))
            d    = self.decl [decl]
            t    = OOo_Tag ('text', decl)
            for n in decls.findall ('.//' + t) :
                name = n.get (_textname)
                if name not in d :
                    sect.append (n)
                    d [name] = 1
    # end def update_decls

    def body_concat (self, root) :
        n = {}
        for i in 'page-count', 'character-count', 'paragraph-count' :
            n [i] = self._get_meta (i)
        pb   = Addpagebreak \
            (stylename = self.pbname, transformer = self.transformer)
        self.divide_body (root)
        self.register_decls ()
        for idx in range (len (self.docs)) :
            meta = self.docs [idx].read ('meta.xml').getroot ()
            self.getmeta.apply (meta)
            ra = Attribute_Access \
                (( Reanchor (n ['page-count'], OOo_Tag ('draw', 'text-box'))
                ,  Reanchor (n ['page-count'], OOo_Tag ('draw', 'rect'))
                ))
            for i in 'page-count', 'character-count', 'paragraph-count' :
                val = self.transformer \
                    [':'.join (('Get_Attribute', 'concat' + i))]
                n [i] += int (val)
            n ['paragraph-count'] += 1
            map = self.stylemaps [idx]
            tree = self.treemaps ['content.xml'][idx]
            r1 = set_attributes_from_dict (None, _stylename,  map)
            r2 = set_attributes_from_dict (None, _parentname, map)
            tr = Attribute_Access (r1 + r2, transformer = self.transformer)
            pb.apply (self.bodyparts [-1])
            tr.apply (tree)
            ra.apply (tree)
            append = tree.find (OOo_Tag ('office', 'body'))
            declarations = self._divide (append)
            self.update_decls   (declarations)
            self.append_to_body (self.copyparts)
        self.append_declarations ()
        self.assemble_body       ()
        for i in 'page-count', 'character-count', 'paragraph-count' :
            self._set_meta (i, n [i])
    # end def body_concat

    def _newname (self, tag, oldname) :
        stylenum = 0
        if (tag, oldname) not in self.stylenames :
            self.stylenames [(tag, oldname)] = 1
            return oldname
        newname = basename = 'Concat_%s' % oldname
        while (tag, newname) in self.stylenames :
            stylenum += 1
            newname = '%s%d' % (basename, stylenum)
        stylenum += 1
        return newname
    # end def _newname

    def style_register (self, oofile) :
        """
            Loop over all style elements in document, serialise them
            and put them in a dict by their name. For each tag register
            the various names.
        """
        for node in self.trees [oofile].getroot () :
            if node.tag in self.style_containers :
                self.containers [oofile][node.tag] = node
                idx       = 0
                to_delete = []
                map       = {}
                for n in node :
                    name  = n.get (_stylename, None)
                    pname = n.get (_parentname, None)
                    pname = map.get (pname, pname)
                    if name :
                        sn  = tree_serialise (n)
                        key = (sn, pname)
                        if key in self.serialised :
                            # looks like OOo has duplicate font definitions
                            # in styles.xml and content.xml -- don't know if
                            # this is needed by OOo or if we could optimize
                            # these away.
                            newname = self.serialised [key]
                            if name != newname :
                                map [name] = newname
                            if node.tag != OOo_Tag ('office', 'font-decls') :
                                to_delete.append (idx)
                        else :
                            self.serialised [key] = name
                            assert (n.tag, name) not in self.stylenames
                            self.stylenames [(n.tag, name)] = 1
                    idx += 1
                to_delete.reverse ()
                for i in to_delete :
                    del node [i]
                if map :
                    r1 = set_attributes_from_dict (None, _stylename,  map)
                    r2 = set_attributes_from_dict (None, _parentname, map)
                    tr = Attribute_Access \
                        (r1 + r2, transformer = self.transformer)
                    tr.apply (self.trees ['content.xml'].getroot ())
                    if self.pbname in map :
                        self.pbname = map [self.pbname]
    # end style_register
            
    def style_concat (self, oofile) :
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
            defined, we insert the parent style into the stylemap dict.
            If the parent style is defined later, it is already too
            late, so an assertion is raised in this case.
        """
        for idx in range (len (self.docs)) :
            stylemap = self.stylemaps [idx]
            root     = self.treemaps [oofile][idx]
            for node in root :
                if node.tag in self.style_containers :
                    for n in node :
                        name  = n.get (_stylename, None)
                        pname = n.get (_parentname, None)
                        pname = stylemap.get (pname, pname)
                        if not name : continue
                        sn  = tree_serialise (n)
                        key = (sn, pname)
                        if key in self.serialised :
                            newname = self.serialised [key]
                            if name != newname :
                                assert \
                                    (  name not in stylemap
                                    or stylemap [name] == newname
                                    )
                                stylemap [name] = newname
                        else :
                            parent = n.get (_parentname, None)
                            if parent :
                                newp = stylemap.get (parent, None)
                                if newp :
                                    n.set (_parentname, newp)
                                else :
                                    # keep name for cross-check
                                    stylemap [_parentname] = newp
                            newname = self._newname (n.tag, name)
                            self.serialised [key] = newname
                            if newname != name :
                                n.set (_stylename, newname)
                                stylemap [name]       = newname
                            self.containers [oofile][node.tag].append (n)
    # end def style_concat

# end class Concatenate

renumber_frames   = Renumber (OOo_Tag ('draw',  'text-box'), 'Frame')
renumber_sections = Renumber (OOo_Tag ('text',  'section'))
renumber_tables   = Renumber (OOo_Tag ('table', 'table'))

# used to have a separate Pagecount transform -- generalized to get
# some of the meta information using an Attribute_Access transform
# and set the same information later after possibly being updated by
# other transforms. We use another naming convention here for storing
# the info retrieve from the OOo document: We use the attribute name in
# the meta-information to store (and later retrieve) the information.

get_attr = []
set_attr = []
for attr in \
    ( 'character-count', 'image-count', 'object-count', 'page-count'
    , 'paragraph-count', 'table-count', 'word-count'
    ) :
    a = OOo_Tag ('meta', attr)
    t = OOo_Tag ('meta', 'document-statistic')
    get_attr.append (Get_Attribute (t, a, attr))
    set_attr.append (Set_Attribute (t, a, attr))
get_meta = Attribute_Access (get_attr, prio =  20, filename = 'meta.xml')
set_meta = Attribute_Access (set_attr, prio = 120, filename = 'meta.xml')
