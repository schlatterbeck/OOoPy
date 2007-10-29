#!/usr/bin/env python

import sys
from ooopy.OOoPy import OOoPy

def as_text (node) :
    if node.text is not None :
        print node.text.encode ('utf-8'),
    for subnode in node :
        as_text (subnode)
    if node.tail is not None :
        print node.tail.encode ('utf-8'),

if __name__ == '__main__' :
    o = OOoPy  (infile = sys.argv [1])
    e = o.read ('content.xml')
    as_text (e.getroot ())
