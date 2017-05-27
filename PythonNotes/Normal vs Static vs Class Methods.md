
### The Normal Method

First I'll explain the normal_method. This may be better called an "instance method". When an instance method is used, it is used as a partial function (as opposed to a total function, defined for all values when viewed in source code) that is, when used, the first of the arguments is predefined as the instance of the object, with all of its given attributes. It has the instance of the object bound to it, and it must be called from an instance of the object. Typically, it will access various attributes of the instance.

For example, this is an instance of a string:

      ', '
if we use the instance method, join on this string, to join another iterable, it quite obviously is a function of the instance, in addition to being a function of the iterable list, ['a', 'b', 'c']:

      ', '.join(['a', 'b', 'c'])
      'a, b, c'

### Static Method

The static method does not take the instance as an argument. Yes it is very similar to a module level function. However, a module level function must live in the module and be specially imported to other places where it is used. If it is attached to the object, however, it will follow the object conveniently through importing and inheritance as well.

An example is the str.maketrans static method, moved from the string module in Python 3. It makes a translation table suitable for consumption by str.translate. It does seem rather silly when used from an instance of a string, as demonstrated below, but importing the function from the string module is rather clumsy, and it's nice to be able to call it from the class, as in str.maketrans

 demonstrate same function whether called from instance or not:
 
      ', '.maketrans('ABC', 'abc')
      {65: 97, 66: 98, 67: 99}
      str.maketrans('ABC', 'abc')
      {65: 97, 66: 98, 67: 99}

In python 2, you have to import this function from the increasingly deprecated string module:

      import string
      'ABCDEFG'.translate(string.maketrans('ABC', 'abc'))
       'abcDEFG'

### Class Method

A class method is a similar to a static method in that it takes an implicit first argument, but instead of taking the instance, it takes the class. Frequently these are used as alternative constructors for better semantic usage and it will support inheritance.

The most canonical example of a builtin classmethod is dict.fromkeys. It is used as an alternative constructor of dict, (well suited for when you know what your keys are and want a default value for them.)

      dict.fromkeys(['a', 'b', 'c'])
        {'c': None, 'b': None, 'a': None}

When we subclass dict, we can use the same constructor, which creates an instance of the subclass.

    class MyDict(dict): 'A dict subclass, use to demo classmethods'
    md = MyDict.fromkeys(['a', 'b', 'c'])
    md
    {'a': None, 'c': None, 'b': None}
    type(md)
    <class '__main__.MyDict'>
