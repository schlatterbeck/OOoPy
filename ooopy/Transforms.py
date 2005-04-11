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
from OOoPy                   import OOoPy
from Transformer             import files, namespaces, OOo_Tag, Transform
from copy                    import deepcopy

def _body () :
    """
        We use the body element as a container for various
        transforms...
    """
    return Element (OOo_Tag ('office', 'body'))
# end def _body

class Renumber (object) :
    """ Specifies a renumbering transform. OOo has a 'name' attribute
        for several different tags, e.g., tables, frames, sections etc.
        These names must be unique in the whole document. OOo itself
        solves this by appending a unique number to a basename for each
        element, e.g., sections are named 'Section1', 'Section2', ...
        Renumber transforms can be applied to correct the numbering
        after operations that destroy the unique numbering, e.g., after
        a mailmerge where the same document is repeatedly appended.

        For performance reasons we do not specify a separate transform
        for each renumbering operation. Instead we define all the
        renumberings we want to perform as Renumber objects that follow
        the attribute changer api and apply them all using an
        Attribute_Changer_Transform in one go.
    """

    def __init__ (self, namespace, tag, name = None, attr = None, start = 1) :
        self.namespace  = namespace
        self.name       = name or tag [0].upper () + tag [1:]
        self.num        = start
        self.tag        = OOo_Tag (namespace, tag)
        self.attribute  = attr or OOo_Tag (namespace, 'name')
    # end def __init__

    def new_value (self, oldval = None) :
        name = "%s%d" % (self.name, self.num)
        self.num += 1
        return name
    # end def new_value

# end class Renumber

class Reanchor (object) :
    """
        Similar to the renumbering transform in that we are assigning
        new values to some attributes. But in this case we want to
        relocate objects that are anchored to a page.
        
        This is another example of a class following the attribute
        changer API.
    """

    def __init__ (self, offset, namespace, tag, attr = None) :
        self.offset     = offset
        self.namespace  = namespace
        self.tag        = OOo_Tag (namespace, tag)
        self.attribute  = attr or OOo_Tag ('text', 'anchor-page-number')
    # end def __init__

    def new_value (self, oldval) :
        if oldval is None :
            return oldval
        return "%d" % (int (oldval) + self.offset)
    # end def new_value

# end class Reanchor

#
# meta.xml transforms
#

class Editinfo_Transform (Transform) :
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
# end class Editinfo_Transform

class Pagecount_Transform (Transform) :
    """
        This is an example of getting information from an OOo document.
        We simply read the number of pages in the document and store it
        for later use by other transforms.
        The stored value can be retrieved from the transformer with
        transformer ['Pagecount_Transform:pagecount'] by other
        transforms.
    """
    filename = 'meta.xml'
    prio     = 20

    def apply (self, root) :
        count = root.find ('.//' + OOo_Tag ('meta', 'document-statistic'))
        self.set ('pagecount', int (count.get (OOo_Tag ('meta', 'page-count'))))
    # end def apply
# end class Pagecount_Transform

#
# settings.xml transforms
#

class Autoupdate_Transform (Transform) :
    """
        This is an example of modifying OOo settings. We set some of the
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

#
# content.xml transforms
#

class Field_Replace_Transform (Transform) :
    """
        Takes a dict of replacement key-value pairs. The key is the name
        of a variable in OOo. Additional replacement key-value pairs may
        be specified in **kw. Alternatively a callback mechanism for
        variable name lookups is provided. The callback function is
        given the name of a variable in OOo and is expected to return
        the replacement value or None if the variable value should not
        be replaced.
    """
    filename = 'content.xml'
    prio     = 100

    def __init__ (self, prio = None, replace = None, **kw) :
        """ replace is something behaving like a dict or something
            callable for name lookups
        """
        Transform (prio)
        self.replace  = replace or {}
        self.dict     = kw
    # end def __init__

    def apply (self, root) :
        body = root
        if body.tag != OOo_Tag ('office', 'body') :
            body = body.find (OOo_Tag ('office', 'body'))
        for node in body.findall ('.//' + OOo_Tag ('text', 'variable-set')) :
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
# end class Field_Replace_Transform

class Addpagebreak_Style_Transform (Transform) :
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
    prio     = 80
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
            , { OOo_Tag ('fo', 'break-before') : 'page' }
            )
        self.set ('stylename', stylename)
    # end def apply
# end class Addpagebreak_Style_Transform

class Addpagebreak_Transform (Transform) :
    """
        This transformation adds a page break to the last page of the OOo
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

    def __init__ (self, prio = None, stylename = None, stylekey = None) :
        Transform.__init__ (self, prio)
        self.stylename = stylename
        self.stylekey  = stylekey or 'Addpagebreak_Style_Transform:stylename'
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
# end class Addpagebreak_Transform

class Mailmerge_Transform (Transform) :
    """
        This transformation is used to create a mailmerge document using
        the current document as the template. In the constructor we get
        an iterator that provides a data set for each item in the
        iteration. Elements the iterator has to provide are either
        something that follows the Mapping Type interface (it looks like
        a dict) or something that is callable and can be used for
        name-value lookups.

        A precondition for this transform is the application of the
        Addpagebreak_Style_Transform to guarantee that we know the style
        for adding a page break to the current document. Alternatively
        the stylename (or the stylekey if a different name should be used
        for lookup in the current transformer) can be given in the
        constructor.
    """
    filename = 'content.xml'
    prio     = 100

    sections = \
        [ { OOo_Tag ('text', 'variable-decls') : 1
          , OOo_Tag ('text', 'sequence-decls') : 1
          }
        , { OOo_Tag ('draw', 'text-box')       : 1
          , OOo_Tag ('draw', 'rect')           : 1
          }
        ]

    def __init__ \
        (self, iterator, prio = None, stylename = None, stylekey = None) :
        Transform.__init__ (self, prio)
        self.iterator  = iterator
        self.stylename = stylename
        self.stylekey  = stylekey
    # end def __init__

    def _divide_body (self) :
        """ Divide self.copy into parts that must keep their sequence.
            We use another body tag for storing the parts...
        """
        self.copyparts = _body ()
        self.copyparts.append (_body ())
        l = len (self.sections)
        idx = 0
        for e in self.copy :
            if idx < l :
                if e.tag not in self.sections [idx] :
                    self.copyparts.append (_body ())
                    idx += 1
            self.copyparts [-1].append (e)
    # end def _divide_body

    def apply (self, root) :
        """
            Copy old body, create new empty one and repeatedly append the
            new body.
        """
        pb         = Addpagebreak_Transform \
            (stylename = self.stylename, stylekey = self.stylekey)
        pb.register (self.transformer)
        pagecount  = self.transformer ['Pagecount_Transform:pagecount']
        ra         = Attribute_Changer_Transform \
            (( Reanchor (pagecount, 'draw', 'text-box')
            ,  Reanchor (pagecount, 'draw', 'rect')
            ))
        cont       = root
        if cont.tag != OOo_Tag ('office', 'document-content') :
            cont   = root.find  (OOo_Tag ('office', 'document-content'))
        body       = cont.find  (OOo_Tag ('office', 'body'))
        idx        = cont [:].index (body)
        self.copy  = cont [idx]
        body       = cont [idx] = _body ()
        self._divide_body ()
        bodyparts  = [_body () for i in self.copyparts]

        for i in self.iterator :
            fr = Field_Replace_Transform (replace = i)
            fr.register (self.transformer)
            # add page break only to non-empty body
            # reanchor only after the first mailmerge
            if body :
                pb.apply (bodyparts [-1])
                ra.apply (self.copyparts)
            else :
                for e in self.copyparts [0] :
                    body.append (e)
                del bodyparts [0]
                del (self.copyparts [0]) # non-repeatable parts
            cp = deepcopy (self.copyparts)
            fr.apply (cp)
            for i in range (len (bodyparts)) :
                for j in cp [i] :
                    bodyparts [i].append (j)
        for p in bodyparts :
            for e in p :
                body.append (e)
    # end def apply
# end class Mailmerge_Transform

class Attribute_Changer_Transform (Transform) :
    """
        Change attributes in an OOo document.
        Can be used for renumbering, moving anchored objects, etc.
        Expects a list of attribute changer objects that follow the
        attribute changer API. This API is very simple:

        - Member function "new_value" returns the new value of an
          attribute
        - The attribute "tag" gives the tag for an element we are
          searching
        - The attribute "attribute" gives the name of the attribute we
          want to change.
        For examples of the attribute changer API, see Renumber and
        Reanchor above.
    """
    filename = 'content.xml'
    prio     = 110

    def __init__ (self, attrchangers, prio = None) :
        Transform.__init__ (self, prio)
        self.attrchangers = {}
        for r in attrchangers :
            self.attrchangers [r.tag] = r
    # end def __init__

    def apply (self, root) :
        """ Search for all tags for which we renumber and replace name """
        for n in root.findall ('.//*') :
            if n.tag in self.attrchangers :
                r    = self.attrchangers [n.tag]
                nval = r.new_value (n.get (r.attribute))
                if nval is not None :
                    n.set (r.attribute, nval)
    # end def apply
# end class Attribute_Changer_Transform

renumber_frames   = Renumber ('draw',  'text-box', 'Frame')
renumber_sections = Renumber ('text',  'section')
renumber_tables   = Renumber ('table', 'table')
