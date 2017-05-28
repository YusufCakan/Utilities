# Object Oriented Programming In Python



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
          
### Multiple Constructors

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



### Working with the Python Super Function
Python 2.2 saw the introduction of a built-in function called “super,” which returns a proxy object to delegate method calls to a class – which can be either parent or sibling in nature.

That description may not make sense unless you have experience working with Python, so we’ll break it down.

Essentially, the super function can be used to gain access to inherited methods – from a parent or sibling class – that has been overwritten in a class object.

Or, as the official Python documentation says:


*“Super is used to] return a proxy object that delegates method calls to a parent or sibling class of type. This is useful for accessing inherited methods that have been overridden in a class. The search order is same as that used by getattr() except that the type itself is skipped.”*


#### How Is the Super Function Used?
The super function is somewhat versatile, and can be used in a couple of ways.

**Use Case 1:** Super can be called upon in a single inheritance, in order to refer to the parent class or multiple classes without explicitly naming them. It’s somewhat of a shortcut, but more importantly, it helps keep your code maintainable for the foreseeable future.

**Use Case 2:** Super can be called upon in a dynamic execution environment for multiple or collaborative inheritance. This use is considered exclusive to Python, because it’s not possible with languages that only support single inheritance or are statically compiled.

When the super function was introduced it sparked a bit of controversy. Many developers found the documentation unclear, and the function itself to be tricky to implement. It even garnered a reputation for being harmful. But it’s important to remember that Python has evolved considerably since 2.2 and many of these concerns no longer apply.

The great thing about super is that it can be used to enhance any module method. Plus, there’s no need to know the details about the base class that’s being used as an extender. The super function handles all of it for you.

So, for all intents and purposes, super is a shortcut to access a base class without having to know its type or name.

In Python 3 and above, the syntax for super is:

     super().methoName(args)

Whereas the normal way to call super (in older builds of Python) is:


    super(subClass, instance).method(args)

As you can see, the newer version of Python makes the syntax a little simpler.



How to Call Super in Python 2 and Python 3?
First, we’ll take a regular class definition and modify it by adding the super function. The initial code will look something like this:

 

     class MyParentClass(object):
          def __init__(self):
               pass
          
     class SubClass(MyParentClass):
          def __init__(self):
               MyParentClass.__init__(self)
 

As you can see, this is a setup commonly used for single inheritance. We can see that there’s a base or parent class (also sometimes called the super class), and a denoted subclass.

But we still need to initialize the parent class within the subclass. To make this process easier, Python’s core development team created the super function. The goal was to provide a much more abstract and portable solution for initializing classes.

If we were using Python 2, we would write the subclass like this (using the super function):

 

     class SubClass(MyParentClass):
          def __init__(self):
               super(SubClass, self).__init__()
 

The same code is slightly different when writing in Python 3, however.

     class MyParentClass():
          def __init__(self):
               pass

     class SubClass(MyParentClass):
          def __init__(self):
               super()
 

Notice how the parent class isn’t directly based on the object base class anymore? In addition, thanks to the super function we don’t need to pass it anything from the parent class. Don’t you agree this is much easier?

Now, keep in mind most classes will also have arguments passed to them. The super function will change even more when that happens.

It will look like the following:

     class MyParentClass():
          def __init__(self, x, y):
               pass

     class SubClass(MyParentClass):
          def __init__(self, x, y):
               super().__init__(x, y)
 
Again, this process is much more straightforward than the traditional method. In this case, we had to call the super function’s __init__ method to pass our arguments.

#### What Is the Super Function for Again?
The super function is extremely useful when you’re concerned about forward compatibility. By adding it to your code, you can ensure that your work will stay operational into the future with only a few changes across the board.

Ultimately, it eliminates the need to declare certain characteristics of a class, provided you use it correctly.

In order to use the function properly, the following conditions must be met:

The method being called upon by super() must exist
Both the caller and callee functions need to have a matching argument signature
Every occurrence of the method must include super() after you use it
You might start with a single inheritance class, but later, if you decide to add another base class – or more – the process goes a lot more smoothly. You only need to make a few changes as opposed to a lot.

There has been talk of using the super function for dependency injection, but we haven’t seen any solid examples of this – at least not practical ones. For now, we’re just going to stick with the description we’ve given.

Either way, now you’ll understand that super isn’t as bad as other devs purport it to be.

