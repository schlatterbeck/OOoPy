from __future__ import print_function
import doctest
import os
import sys

if sys.version.startswith ('2.3') :
    # Doctest has a bug with super in python 2.3. This patches the 2.3
    # doctest according to the patch by Christian Tanzer on sourceforge
    # -- the fix is already in 2.4 but according to the docs on
    # sourceforge will not be backported to 2.3. See
    # http://sf.net/tracker/?func=detail&atid=105470&aid=902628&group_id=5470
    # http://mail.python.org/pipermail/python-bugs-list/2004-July/024211.html
    class Tester_patched (doctest.Tester) :
        __record_outcome = doctest.Tester._Tester__record_outcome
        def rundoc (self, obj, name = None) :
            if name is None:
                try:
                    name = obj.__name__
                except AttributeError :
                    raise ValueError("Tester.rundoc: name must be given "
                        "when obj.__name__ doesn't exist; " + repr (obj))
            if self.verbose:
                print ("Running", name + ".__doc__")
            f, t = doctest.run_docstring_examples \
                ( obj, self.globs, self.verbose, name
                , self.compileflags, self.optionflags
                )
            if self.verbose:
                print (f, "of", t, "examples failed in", name + ".__doc__")
            self.__record_outcome(name, f, t)
            if doctest._isclass(obj):
                # In 2.2, class and static methods complicate life.  Build
                # a dict "that works", by hook or by crook.
                d = {}
                for tag, kind, homecls, value in doctest._classify_class_attrs(obj):

                    if homecls is not obj:
                        # Only look at names defined immediately by the class.
                        continue

                    elif self.isprivate(name, tag):
                        continue

                    elif kind == "method":
                        if not isinstance(value, super):
                            # value is already a function
                            d[tag] = value

                    elif kind == "static method":
                        # value isn't a function, but getattr reveals one
                        d[tag] = getattr(obj, tag)

                    elif kind == "class method":
                        d[tag] = getattr(obj, tag).im_func

                    elif kind == "property":
                        # The methods implementing the property have their
                        # own docstrings -- but the property may have one too.
                        if value.__doc__ is not None:
                            d[tag] = str(value.__doc__)

                    elif kind == "data":
                        # Grab nested classes.
                        if doctest._isclass(value):
                            d[tag] = value

                    else:
                        raise ValueError("teach doctest about %r" % kind)
                f2, t2 = self.run__test__(d, name)
                f += f2
                t += t2

            return f, t
    doctest.Tester_orig = doctest.Tester
    doctest.Tester      = Tester_patched

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

