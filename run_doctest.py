from __future__ import print_function
import doctest
import os
import sys

format = "%(file)s fails %(f)s of %(t)s doc-tests"
for a in sys.argv [1:] :
    sys.path [0:0] = ["./", os.path.dirname  (a)]
    os.environ ['PYTHONPATH'] = ':'.join (sys.path)
    m = os.path.splitext (os.path.basename (a)) [0]
    try :
        module = __import__ (m)
        file   = module.__file__
        f, t   = doctest.testmod (module, verbose = 0)
    except KeyboardInterrupt :
        raise
    except Exception as cause :
        print ("Testing of %s resulted in exception" % (a,))
        raise
    else :
        print (format % locals ())
    del sys.path [0:2]

