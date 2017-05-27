
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
       
#### Summary
      1. Static methods are simply functions. They do not have their first argument as self.
      2.. We can acess static methods without initialising the object. Therefore the class that it 
      is held within becomes just a namespace. We do not need to create on object first in order 
      to access its methods.
      3. Static members cannot call self method but can call other static methods and class methods.
      4. Static methods belong to a class and not to the object at all. It works on class attributes 
      and not on instance attributes.
      5.  Static methoids  can be called by both class and instances
      6.  Main Benefits is that it localises function name in the class scope. Sort of allowing functional
      programming pradigm.
      7. Second benefit is that it moves code closer to where it will be used.


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


#### Summary
1. To decide whether to use @staticmethod or @classmethod you have to look inside your method. 
   If your method accesses other variables/methods in your class then use @classmethod.
2. Class methods are not bound to the object but to the class.
3. They can be called without instiating the class.
4. Its first argument must be a class. This means you can use the class and its properties 
inside that method rather than a particular instance.
5. It is useful for:
* Factory methods, that are used to create an instance for a class using for example some sort of pre-processing.
* Static methods calling static methods: if you split a static methods in several static
methods, you shouldn't hard-code the class name but use class methods
            
            
            
            
