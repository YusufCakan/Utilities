
    class Cheese(object):
        def __init__(self, num_holes=0):
            "defaults to a solid cheese"
            self.number_of_holes = num_holes

        @classmethod
        def random(cls):
            return cls(randint(0, 100))

        @classmethod
        def slightly_holey(cls):
            return cls(randint((0,33))

        @classmethod
        def very_holey(cls):
            return cls(randint(66, 100))
     
Now create object like this:

    gouda = Cheese()
    emmentaler = Cheese.random()
    leerdammer = Cheese.slightly_holey()


A much neater way to get 'alternate constructors' is to use classmethods. For instance:

    class MyData:
         def __init__(self, data):
             "Initialize MyData from a sequence"
             self.data = data
     
         @classmethod
         def fromfilename(cls, filename):
             "Initialize MyData from a file"
             data = open(filename).readlines()
             return cls(data)
     
         @classmethod
         def fromdict(cls, datadict):
             "Initialize MyData from a dict's items"
             return cls(datadict.items())
 
         MyData([1, 2, 3]).data
            [1, 2, 3]
         MyData.fromfilename("/tmp/foobar").data
            ['foo\n', 'bar\n', 'baz\n']
         MyData.fromdict({"spam": "ham"}).data
            [('spam', 'ham')]

The reason it's neater is that there is no doubt about what type is expected, and you aren't forced to guess at what the caller intended for you to do with the datatype it gave you. The problem with isinstance(x, basestring) is that there is no way for the caller to tell you, for instance, that even though the type is not a basestring, you should treat it as a string (and not another sequence.) And perhaps the caller would like to use the same type for different purposes, sometimes as a single item, and sometimes as a sequence of items. Being explicit takes all doubt away and leads to more robust and clearer code.
